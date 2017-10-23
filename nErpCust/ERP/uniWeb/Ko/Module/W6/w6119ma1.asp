
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �������鼼������ 
'*  3. Program ID           : W1111MA1
'*  4. Program Name         : W1111MA1.asp
'*  5. Program Desc         : ��8ȣ �������鼼�װ�꼭(4)
'*  6. Modified date(First) : 2005/03/14
'*  7. Modified date(Last)  : 2005/03/14
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

Const BIZ_MNU_ID	 = "W6119MA1"
Const BIZ_PGM_ID	 = "W6119mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID = "W6119mb2.asp"
Const BIZ_POP_ID	 = "W6119MA2.ASP"
Const EBR_RPT_ID	 = "W6119OA1"

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.
Const TYPE_3	= 2		'

Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8

Dim C_W_TYPE
Dim C_SEQ_NO
Dim C_W9
Dim C_W9_NM
Dim C_W10
Dim C_W10_NM
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14
Dim C_W15
Dim C_W16
Dim C_W17
Dim C_W18
Dim C_W18_VIEW
Dim C_W18_VAL
Dim C_W19
Dim C_W20
Dim C_W21
Dim C_W35

Dim IsOpenPop          
Dim lgStrPrevKey2, IsRunEvents
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2)
Dim lgFISC_START_DT, lgFISC_END_DT 
'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W1		= 0	' HTML �ε��� 
	C_W2		= 1
	C_W3		= 2
	C_W4		= 3
	C_W5		= 4
	C_W6		= 5
	C_W7		= 6
	C_W8		= 7

	C_W_TYPE	= 1	' �׸��� �ε��� 
	C_SEQ_NO	= 2
	C_W9		= 3
	C_W9_NM		= 4
	C_W10		= 5
	C_W10_NM	= 6
	C_W11		= 7
	C_W12		= 8
	C_W13		= 9
	C_W14		= 10
	C_W15		= 11
	C_W16		= 12
	C_W17		= 13
	C_W18		= 14
	C_W18_VIEW	= 15
	C_W18_VAL	= 16
	C_W19		= 17
	C_W20		= 18
	C_W21		= 19
	C_W35		= 20
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
    IsRunEvents = False
    
    lgCurrGrid = TYPE_2
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
	Dim ret, iType, i
	
    Call initSpreadPosVariables()  

	Set lgvspdData(TYPE_1) = frm1.txtData
	Set lgvspdData(TYPE_2) = frm1.vspdData0
	Set lgvspdData(TYPE_3) = frm1.vspdData1

	Call AppendNumberPlace("6","5","2")

	' 1�� �׸��� 
	For iType = TYPE_2 To TYPE_3
			
		With lgvspdData(iType)
			
		ggoSpread.Source = lgvspdData(iType)	
		'patch version
		ggoSpread.Spreadinit "V20041222" & iType,,parent.gAllowDragDropSpread    
			 
		.ReDraw = false

		.MaxCols = C_W35 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
			 
  		'����� 3�ٷ�    
		.ColHeaderRows = 2  
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W_TYPE,	"����"		, 10,,,10,1	' �����÷� 
		ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,15,1	' �����÷� 
		ggoSpread.SSSetCombo	C_W9,		"(9)����Ƚ��"		, 10
		ggoSpread.SSSetCombo	C_W9_NM,	"(9)����Ƚ��"		, 10	
		ggoSpread.SSSetCombo	C_W10,		"(10)����"		, 10
		ggoSpread.SSSetCombo	C_W10_NM,	"(10)����"		, 10	
		ggoSpread.SSSetDate		C_W11,		"(11)����" & vbCrLf & "���" & vbCrLf & "����",			10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetDate		C_W12,		"(12)���" & vbCrLf & "����",			10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetFloat	C_W13,		"(13)�Ѿ�" , 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
		ggoSpread.SSSetFloat	C_W14,		"(14)�ܱ�" & vbCrLf & "���ڰ�" & vbCrLf & "�ں���" , 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
		ggoSpread.SSSetFloat	C_W15,		"(15)����" & vbCrLf & "����" & vbCrLf & "�ں���" , 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
		ggoSpread.SSSetDate		C_W16,	"(16)100%" & vbCrLf & "������" & vbCrLf & "(From)(To)" ,			10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetDate		C_W17,	"(17)50%" & vbCrLf & "������" & vbCrLf & "(From)(To)" ,			10,		2,		Parent.gDateFormat,	-1
		ggoSpread.SSSetCombo	C_W18,		"(18)"		, 10
		ggoSpread.SSSetCombo	C_W18_VIEW,	"(18)����" & vbCrLf & "�������" & vbCrLf & "������" & vbCrLf & "(16�Ǵ�17)"		, 10	
		ggoSpread.SSSetCombo	C_W18_VAL,	"(18)"		, 10
		ggoSpread.SSSetFloat	C_W19,		"(19)������" & vbCrLf & "�ܱ����ڰ�" & vbCrLf & "�ں���(14-15*18)" , 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
		ggoSpread.SSSetFloat	C_W20,		"(20)���ػ��" & vbCrLf & "�������ں���" , 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
		ggoSpread.SSSetFloat	C_W21,		"(21)�������" & vbCrLf & "(19/20)" , 12,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
		ggoSpread.SSSetFloat	C_W35,		"(21)�������" & vbCrLf & "(19/20)" , 12,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
		
		ret = .AddCellSpan(C_W_TYPE	, -1000, 1, 2)
		ret = .AddCellSpan(C_SEQ_NO	, -1000, 1, 2)
		ret = .AddCellSpan(C_W9		, -1000, 1, 2)
		ret = .AddCellSpan(C_W9_NM	, -1000, 1, 2)
		ret = .AddCellSpan(C_W10	, -1000, 1, 2)
		ret = .AddCellSpan(C_W10_NM	, -1000, 1, 2)
		ret = .AddCellSpan(C_W11	, -1000, 1, 2) 
		ret = .AddCellSpan(C_W12	, -1000, 1, 2) 
		ret = .AddCellSpan(C_W13	, -1000, 3, 1) 
		ret = .AddCellSpan(C_W16	, -1000, 2, 1) 
		ret = .AddCellSpan(C_W18	, -1000, 1, 2)
		ret = .AddCellSpan(C_W18_VIEW	, -1000, 1, 2)
		ret = .AddCellSpan(C_W18_VAL	, -1000, 1, 2)
		ret = .AddCellSpan(C_W19	, -1000, 1, 2)
		ret = .AddCellSpan(C_W20	, -1000, 1, 2)
		If iType = TYPE_2 Then
			ret = .AddCellSpan(C_W21	, -1000, 2, 2)
			i = 9
		Else
			ret = .AddCellSpan(C_W21	, -1000, 1, 2)
			ret = .AddCellSpan(C_W35	, -1000, 1, 2)
			i = 22
		End If
		
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W9_NM	: .Text = "(" & i & ")����Ƚ��"
		.Col = C_W10_NM	: .Text = "(" & i+1 & ")����"
		.Col = C_W11	: .Text = "(" & i+2 & ")����" & vbCrLf & "���" & vbCrLf & "����"
		.Col = C_W12	: .Text = "(" & i+3 & ")���" & vbCrLf & "����"
		.Col = C_W13	: .Text = "�����ں���" & vbCrLf & "(�պ����ں���)"
		.Col = C_W16	: .Text = "����Ⱓ"
		
		.Col = C_W18_VIEW	: .Text = "(" & i+9 & ")����" & vbCrLf & "�������" & vbCrLf & "������" & vbCrLf & "(" & i+7 & "�Ǵ�" & i+8 &")"
		.Col = C_W19	: .Text = "(" & i+10 & ")������" & vbCrLf & "�ܱ����ڰ�" & vbCrLf & "�ں���" & vbCrLf & "(" & i+5 & "-" & i+6 & "*" & i+9 & ")"
		
		.Col = C_W20	: .Text = "(" & i+11 & ")���ػ��" & vbCrLf & "�������ں���"
		
		If iType = TYPE_2 Then
			.Col = C_W21	: .Text = "(" & i+12 & ")�������" & vbCrLf & "(19/20)"
		Else
			.Col = C_W21	: .Text = "(" & i+12 & ")�ܱ������ں���"
			.Col = C_W35	: .Text = "(" & i+13 & ")�������" & vbCrLf & "(32/33)*34"
		End If
		
		' �ι�° ��� ��� ���� 
		.Row = -999
		.Col = C_W13	: .Text = "(" & i+4 & ")�Ѿ�"
		.Col = C_W14	: .Text = "(" & i+5 & ")�ܱ�" & vbCrLf & "���ڰ�" & vbCrLf & "�ں���"
		.Col = C_W15	: .Text = "(" & i+6 & ")����" & vbCrLf & "����" & vbCrLf & "�ں���"
		.Col = C_W16	: .Text = "(" & i+7 & ")100%" & vbCrLf & "������" & vbCrLf & "(From)(To)"
		.Col = C_W17	: .Text = "(" & i+8 & ")50%" & vbCrLf & "������" & vbCrLf & "(From)(To)"

		.rowheight(-1000) = 20	' ���� ������ 
		.rowheight(-999) = 30	' ���� ������ 
	
		Call ggoSpread.SSSetColHidden(C_W9,C_W9,True)
		Call ggoSpread.SSSetColHidden(C_W10,C_W10,True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W18,C_W18,True)
		Call ggoSpread.SSSetColHidden(C_W18_VAL,C_W18_VAL,True)
							
		Call InitSpreadComboBox(iType)

		.ReDraw = true
	
		End With
     Next  
     
     Call MakePercentCol( lgvspdData(TYPE_2), C_W21, "", "", "")
     Call MakePercentCol( lgvspdData(TYPE_3), C_W21, "", "", "")
     Call MakePercentCol( lgvspdData(TYPE_3), C_W35, "", "", "")
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox(Byval pType)

    Dim IntRetCD1, sVal

	' ȸ�籸�� 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "TB_MINOR", " MAJOR_CD='W1028' AND REVISION_YM='" & C_REVISION_YM & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(pType)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W9
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W9_NM
	End If

	' ȸ�籸�� 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " (MAJOR_CD='W1029') ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(pType)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W10
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W10_NM
	End If


    ggoSpread.Source = lgvspdData(pType)
    sVal = "1" & vbTab & "2" & vbTab & "3"
	ggoSpread.SetCombo sVal, C_W18

    sVal = "100%" & vbTab & "50%" & vbTab & "0%(����Ⱓ����)"
	ggoSpread.SetCombo sVal, C_W18_VIEW
	
	sVal = "1" & vbTab & "0.5" & vbTab & "0"
	ggoSpread.SetCombo sVal, C_W18_VAL
	
End Sub

Sub SetSpreadLock()

    lgvspdData(TYPE_2).ReDraw = False
    ggoSpread.Source = lgvspdData(TYPE_2)	

    ggoSpread.SpreadLock C_W_TYPE, -1, C_SEQ_NO
    
    If lgvspdData(TYPE_2).MaxRows > 1 Then
		ggoSpread.SpreadLock C_W19, 1, C_W21, lgvspdData(TYPE_2).MaxRows-1
	End If
    
	'ggoSpread.SSSetRequired C_W14, -1, -1

	ggoSpread.SpreadLock C_W9, lgvspdData(TYPE_2).MaxRows, C_W19, lgvspdData(TYPE_2).MaxRows
	ggoSpread.SpreadLock C_W21, lgvspdData(TYPE_2).MaxRows, C_W21, lgvspdData(TYPE_2).MaxRows
	lgvspdData(TYPE_2).ReDraw = True
	
	lgvspdData(TYPE_3).ReDraw = False
    ggoSpread.Source = lgvspdData(TYPE_3)
        
    ggoSpread.SpreadLock C_W_TYPE, -1, C_SEQ_NO
    
    If lgvspdData(TYPE_3).MaxRows > 1 Then
		ggoSpread.SpreadLock C_W19, 1, C_W21, lgvspdData(TYPE_3).MaxRows-1
	End If
    
	'ggoSpread.SSSetRequired C_W14, -1, -1

	ggoSpread.SpreadLock C_W9, lgvspdData(TYPE_3).MaxRows, C_W19, lgvspdData(TYPE_3).MaxRows
	ggoSpread.SpreadLock C_W35, lgvspdData(TYPE_3).MaxRows, C_W35, lgvspdData(TYPE_3).MaxRows
	lgvspdData(TYPE_3).ReDraw = True
	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
    With lgvspdData(pType)

	.ReDraw = False
 
	ggoSpread.Source = lgvspdData(pType)
	
	ggoSpread.SSSetProtected C_W_TYPE, pvStartRow, pvEndRow
  	ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow

	Select Case pType
		Case TYPE_2
			ggoSpread.SpreadLock C_W19, pvStartRow, C_W21, pvEndRow
		Case TYPE_3
			ggoSpread.SpreadLock C_W19, pvStartRow, C_W35, pvEndRow
	End Select

  	'ggoSpread.SSSetRequired C_W18_VIEW, pvStartRow, pvEndRow
  				    
	.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
  
End Sub

'============================== ���۷��� �Լ�  ========================================

Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, i
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

	For i = C_W3 To C_W8
		frm1.txtData(i).value = ""
	Next
	
	IntRetCD = CommonQueryRs("W01"," TB_3 " , " CO_CD='" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	If IntRetCD = False Then
		Call  DisplayMsgBox("X", parent.VB_INFORMATION, "���μ� ���� ��3ȣ ���� ���μ�����ǥ�� �� ����������꼭�� ����ǥ��(�ڵ�56)/ ���⼼�� �հ�(�ڵ�16) �ݾ��� ���� �����Ͻʽÿ�", "X")
		Exit Function
	End If
	
	' �������� ��� 
	IntRetCD = CommonQueryRs("W4, W5"," dbo.ufn_TB_8_4_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If IntRetCD = True Then
		With frm1
		
		If UNICDbl(Replace(lgF0, Chr(11), "")) = 0 Then
			Call DisplayMsgBox("W60006", parent.VB_INFORMATION, "����ǥ�� �ݾ�", "X")
			Exit Function
		End If
		IsRunEvents = True
		.txtData(C_W4).value = Replace(lgF0, Chr(11), "")
		.txtData(C_W5).value = Replace(lgF1, Chr(11), "")		
		IsRunEvents = False
		End With
    End If
	
	Call SetHeadReCalc()
	
	lgBlnFlgChgValue = True
End Function

Sub PutGrid(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	If pType = TYPE_1 Then
		With lgvspdData(TYPE_1)
			.Col = pCol	: .Row = pRow : .Value = pVal
		End With
	Else
		With lgvspdData(TYPE_2)
			.Col = pCol	: .Row = pRow : .Value = pVal
		End With
	End If
End Sub

Sub PutGridText(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	If pType = TYPE_1 Then
		With lgvspdData(TYPE_1)
			.Col = pCol	: .Row = pRow : .Text = pVal
		End With
	Else
		With lgvspdData(TYPE_2)
			.Col = pCol	: .Row = pRow : .Text = pVal
		End With
	End If
End Sub

' -- ��系���ҷ����� ���� ȣ��� 
Sub ReClacGrid2()
	Dim dblVal(30), iMaxRows, iRow
	
	With lgvspdData(TYPE_2)
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			
			.Col = C_W15	: dblVal(C_W15) = UNICDbl(.value)
			.Col = C_W16	: dblVal(C_W16) = UNICDbl(.Text)
			If dblVal(C_W15) = 0 And dblVal(C_W16) = 0 Then Exit Sub
			dblVal(C_W17) = dblVal(C_W15) * (dblVal(C_W16) * 0.01)
			.Col = C_W17	: .value = dblVal(C_W17)		
		
		Next
	
	
	End With
End Sub

' -- 1 html ���� ��� 
Sub SetHeadReCalc()
	If IsRunEvents = True Then Exit Sub	' -- �̺�Ʈ �湮�� ������� 
	
	IsRunEvents = True
	
	Dim dblAmt(8)
	
	With frm1

		dblAmt(C_W2) = UNICDbl(.txtData(C_W2).value)
		dblAmt(C_W4) = UNICDbl(.txtData(C_W4).value)
		dblAmt(C_W5) = UNICDbl(.txtData(C_W5).value)
		If dblAmt(C_W4) = 0 Then 
			IsRunEvents = False
			Exit Sub
		End If
		dblAmt(C_W3) = dblAmt(C_W4) - dblAmt(C_W2)
		If dblAmt(C_W3) < 0 Then
			dblAmt(C_W3) = 0
		End If
		.txtData(C_W3).value = dblAmt(C_W3)
		
		dblAmt(C_W6) = formatNumber((dblAmt(C_W2) / dblAmt(C_W4)) * 100, 2)
		.txtData(C_W6).value = dblAmt(C_W6) & "%"
		
		' W7 = ??
		If lgvspdData(TYPE_2).MaxRows > 0 Then
			dblAmt(C_W7)			= GetGrid(TYPE_2, C_W21, lgvspdData(TYPE_2).MaxRows)
			.txtData(C_W7).value	= GetGridText(TYPE_2, C_W21, lgvspdData(TYPE_2).MaxRows)
		ElseIf lgvspdData(TYPE_3).MaxRows > 0 Then
			dblAmt(C_W7)			= GetGrid(TYPE_3, C_W35, lgvspdData(TYPE_3).MaxRows)
			.txtData(C_W7).value = GetGridText(TYPE_3, C_W35, lgvspdData(TYPE_3).MaxRows)
		Else
			dblAmt(C_W7) = 0
		End If
		
		dblAmt(C_W8) = dblAmt(C_W5) * (dblAmt(C_W6)/100) * dblAmt(C_W7)
		.txtData(C_W8).value = dblAmt(C_W8)
		
	End With
	
	IsRunEvents = False
    lgBlnFlgChgValue = True
End Sub

Function OpenW2()	'�������� 
	Dim arrRet, strWhere
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtCO_CD.Value
	arrParam(1) = frm1.txtFISC_YEAR.Text		
	arrParam(2) = frm1.cboREP_TYPE.Value		

    arrRet = window.showModalDialog(BIZ_POP_ID, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		
	    Exit Function
	Else	
		IsRunEvents = True
		frm1.txtData(C_W2).value = arrRet(0)
		IsRunEvents = False
		Call SetHeadReCalc
	End If
End Function

Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, iCnt
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
		
End Sub

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
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
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

Sub InitData()
	Dim sCoCd , sFiscYear, sRepType
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	lgCurrGrid = TYPE_2
	
	Call GetFISC_DATE
End Sub

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	Call GetFISC_DATE
End Sub

Sub cboW2_onChange()
	With frm1
		.txtW2_NM.value =	.cboW2.options(.cboW2.selectedIndex).Text
	End With
	lgBlnFlgChgValue = True
End Sub

Sub txtW1_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW3_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW4_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW5_onChange()
	lgBlnFlgChgValue = True
End Sub
Sub txtW6_onChange()
	lgBlnFlgChgValue = True
End Sub

'============================================  �׸��� �̺�Ʈ   ====================================
' -- 0�� �׸��� 
Sub vspdData0_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_2
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_2
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData0_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_2
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData0_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_2
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData0_GotFocus()
	lgCurrGrid = TYPE_2
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData0_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_2
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData0_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_2
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData0_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_2
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 1�� �׸��� 
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_ComboSelChange(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_3
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_3
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_3
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_3
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData1_GotFocus()
	lgCurrGrid = TYPE_3
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData1_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_3
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_3
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

Sub vspdData0_ComboCloseUp(ByVal Col , ByVal Row , ByVal SelChange )
	Call vspdData_ComboCloseUp(TYPE_2, Col, Row, SelChange)
End Sub

Sub vspdData1_ComboCloseUp(ByVal Col , ByVal Row , ByVal SelChange )
	Call vspdData_ComboCloseUp(TYPE_3, Col, Row, SelChange)
End Sub

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Byval pType, ByVal Col, ByVal Row)
	Dim iIdx, dblVal(30)
	
	lgBlnFlgChgValue = True
	' -- 2006.02.22(choe0tae) �ش� �̺�Ʈ�� Span�� �޺������� ���׹߻��Ǿ� 2��°���ҵ� �࿡���� �޺��ڽ� ������ ó������ ���� ComboCloseUp ���� �̵��ƴ�.
End Sub

Sub vspdData_ComboCloseUp(Byval pType, ByVal Col , ByVal Row , ByVal SelChange )
	Dim iIdx
	With lgvspdData(pType)
		lgBlnFlgChgValue = True
		
		Select Case Col

			Case C_W9_NM, C_W10_NM
				
				.Row = Row
				.Col = Col
				iIdx = UNICDbl(.Value)
				
				.Col = Col -1 
				.Value = iIdx
		
			Case C_W18_VIEW
	
				.Col = Col
				.Row = Row
				iIdx = UNICDbl(.Value)
				
				.Row = Row
				.Col = Col -1 
				.Value = iIdx
				.Col = Col +1 
				.Value = iIdx
				
				Call vspdData_Change(pType, C_W18_VIEW, Row)

		End Select
    End With
End Sub

Private Function GetGridValue2(ByVal pCol , ByVal pRow )
    With lgvspdData(lgCurrGrid)
        .Col = pCol: .Row = pRow
        GetGridValue2 = .Value
    End With
End Function

Private Function GetGridText2(ByVal pCol , ByVal pRow )
    With lgvspdData(lgCurrGrid)
        .Col = pCol: .Row = pRow
        GetGridText2 = .Text
    End With
End Function
	
Sub vspdData_Change(Byval pType, ByVal Col , ByVal Row )
    lgvspdData(pType).Row = Row
    lgvspdData(pType).Col = Col

    If lgvspdData(pType).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(pType).text) < UNICDbl(lgvspdData(pType).TypeFloatMin) Then
         lgvspdData(pType).text = lgvspdData(pType).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(pType)
    ggoSpread.UpdateRow Row
    
	Dim dblVal(30), dblTerm, iRow
	
	With lgvspdData(pType)

		If Row <> .MaxRows Then
			If (Row Mod 2) = 0 Then
				ggoSpread.UpdateRow Row-1
			Else
				ggoSpread.UpdateRow Row+1
			End If
			ggoSpread.UpdateRow .MaxRows
		End If
	
		Select Case Col
			
			Case C_W13, C_W19
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				
			Case C_W14, C_W15, C_W18_VIEW	' -- W19 ��� 
				.Row = Row
				dblVal(C_W11) = GetGridText(pType, C_W11, Row)
				dblVal(C_W14) = GetGrid(pType, C_W14, Row)
				dblVal(C_W15) = GetGrid(pType, C_W15, Row)
				dblVal(C_W18) = GetGridText(pType, C_W18_VAL, Row)
				
				If dblVal(C_W18) = "" Then dblVal(C_W18) = 0
				
				If dblVal(C_W11) >= lgFISC_START_DT And  dblVal(C_W11)  <= lgFISC_END_DT Then
					dblVal(C_W19) = ( dblVal(C_W14) - dblVal(C_W15) ) * dblVal(C_W18) * ( DateDiff("d", dblVal(C_W11), lgFISC_END_DT)+1 / DateDiff("d", lgFISC_START_DT, lgFISC_END_DT)+1 )
				Else
					dblVal(C_W19) = ( dblVal(C_W14) - dblVal(C_W15) ) * dblVal(C_W18)
				End If
				
				Call PutGrid(pType, C_W19, Row, dblVal(C_W19))
								
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W14, 1, .MaxRows - 1, true, .MaxRows, C_W14, "V")	' �հ� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' �հ� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W19, 1, .MaxRows - 1, true, .MaxRows, C_W19, "V")	' �հ� 
			
				' C_W19 �� ���������Ƿ� �ش��̺�Ʈ(C_W20)�� �߻��� �ش�.
				Call vspdData_Change(pType, C_W19, Row)
				Call vspdData_Change(pType, C_W20, Row)
				
			Case C_W20, C_W21
				iRow = .MaxRows
				.Row = iRow
				dblVal(C_W19) = UNICDbl(GetGrid(pType, C_W19, iRow))
				dblVal(C_W20) = UNICDbl(GetGrid(pType, C_W20, iRow))
				If dblVal(C_W20) = 0 Then Exit Sub
				
				If pType = TYPE_2 Then
					dblVal(C_W21) = dblVal(C_W19) / dblVal(C_W20)
					Call PutGrid(pType, C_W21, iRow, dblVal(C_W21))
					
					Call SetHeadReCalc
				Else
					dblVal(C_W21) = UNICDbl(GetGrid(pType, C_W21, iRow))
					dblVal(C_W35) = ( dblVal(C_W19) / dblVal(C_W20) ) * dblVal(C_W21) 
					Call PutGrid(pType, C_W35, iRow, dblVal(C_W35))
					
					Call SetHeadReCalc
				End If
		End Select
	End With
End Sub

Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol : .Row = pRow : GetGrid = .value
	End With
End Function

Function GetGridText(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol : .Row = pRow : GetGridText = .Text
	End With
End Function

Function PutGrid(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(pType)
		.Col = pCol : .Row = pRow : .value = pVal
	End With
End Function

Function PutGridText(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(pType)
		.Col = pCol : .Row = pRow : .text = pVal
	End With
End Function

Sub vspdData_Click(Byval pType, ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("0001011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = lgvspdData(pType)
   
    If lgvspdData(pType).MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(pType)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	lgvspdData(TYPE_2).Row = Row
End Sub

Sub vspdData_ColWidthChange(Byval pType, ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = lgvspdData(pType)
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(Byval pType, ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If lgvspdData(pType).MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData_GotFocus(Byval pType )
    ggoSpread.Source = lgvspdData(pType)
End Sub

Sub vspdData_MouseDown(Byval pType, Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	ggoSpread.Source = lgvspdData(pType)
End Sub    

Sub vspdData_ScriptDragDropBlock(Byval pType, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = lgvspdData(TYPE_1)
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos(pType)
End Sub

Sub vspdData_TopLeftChange(Byval pType, ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if lgvspdData(pType).MaxRows < NewTop + VisibleRowCnt(lgvspdData(pType),NewTop) Then	           
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
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	If lgBlnFlgChgValue Then
    'ggoSpread.Source = lgvspdData(TYPE_2)
    'If ggoSpread.SSCheckChange = True Then
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
    
    Dim Strchange1,Strchange2
    Dim blnChange
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    

<%  '-----------------------
    'Precheck area
    '----------------------- %> 
    ggoSpread.Source = lgvspdData(TYPE_2)
    Strchange1 = ggoSpread.SSCheckChange
    If ggoSpread.SSCheckChange = False Then
        blnChange = False
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If    
	
    ggoSpread.Source = lgvspdData(TYPE_3)
    Strchange2 = ggoSpread.SSCheckChange
    If ggoSpread.SSCheckChange = False Then
        blnChange = False
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If  

    If lgBlnFlgChgValue = False And Strchange1= False And Strchange2 = False  Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
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

    Call SetToolbar("1100110100000111")

	Call InitData
	
	'frm1.txtCO_CD.focus
    Dim IntRetCD1, sVal

	' ȸ�籸�� 
	IntRetCD1 = CommonQueryRs("W56, W16", "TB_3", "CO_CD=" & FilterVar(Trim(UCase(frm1.txtCO_CD.value)),"''","S") & " AND FISC_YEAR=" & FilterVar(Trim(UCase(frm1.txtFISC_YEAR.text)),"''","S") & " AND REP_TYPE=" & FilterVar(Trim(UCase(frm1.cboREP_TYPE.value)),"''","S") , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 = False Then
		Call  DisplayMsgBox("X", parent.VB_INFORMATION, "���μ� ���� ��3ȣ ���� ���μ�����ǥ�� �� ����������꼭�� ����ǥ��(�ڵ�56)/ ���⼼�� �հ�(�ڵ�16) �ݾ��� ���� �����Ͻʽÿ�", "X")
	End If

    FncNew = True

End Function


Function FncCopy() 

End Function

Function FncCancel() 
	Dim iRow
	
	With lgvspdData(lgCurrGrid)
	ggoSpread.Source = lgvspdData(lgCurrGrid)	
	If .MaxRows = 3 Then

		ggoSpread.EditUndo                                                  '��: Protect system from crashing
		ggoSpread.EditUndo                                                  '��: Protect system from crashing
		ggoSpread.EditUndo                                                  '��: Protect system from crashing
		
	Else
		If (.ActiveRow Mod 2) = 0 Then
			iRow = .ActiveRow - 1
			.Row = iRow : .SetActiveCell .ActiveCol, iRow
			ggoSpread.EditUndo 
			.Row = iRow : .SetActiveCell .ActiveCol, iRow
			ggoSpread.EditUndo 
		Else
			iRow = .ActiveRow
			ggoSpread.EditUndo 
			.Row = iRow : .SetActiveCell .ActiveCol, iRow
			ggoSpread.EditUndo 
		End If

		ggoSpread.UpdateRow .MaxRows

		Call vspdData_Change(TYPE_2, C_W13, .ActiveRow)
		Call vspdData_Change(TYPE_2, C_W14, .ActiveRow)
		'Call vspdData_Change(TYPE_2, C_W20, .ActiveRow)

	End If
			
	
	End With
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID, blnError

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
	
	blnError = False
	imRow = imRow * 2
	
	With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 
	
		Select Case lgCurrGrid
			Case TYPE_2
				If lgvspdData(TYPE_3).MaxRows > 0 Then
					Call DisplayMsgBox("X", parent.VB_INFORMATION, "3. �񰨸����� �����ϴ�... �׸��忡 ����Ÿ�� ������ ��쿡�� 2. �Ϲ����� ������� ��� �׸��忡 ���߰��� �� �� �����ϴ�.", "X")
					Exit Function
				End If
			Case TYPE_3
				If lgvspdData(TYPE_2).MaxRows > 0 Then
					Call DisplayMsgBox("X", parent.VB_INFORMATION, "2. �Ϲ����� ������� ��� �׸��忡 ����Ÿ�� ������ ��쿡�� 3. �񰨸����� �����ϴ�... �׸��忡 ���߰��� �� �� �����ϴ�.", "X")
					Exit Function
				End If
		End Select
		
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		.Row = .ActiveRow
		iRow = .ActiveRow	: If iRow < 1 Then iRow = 1
		lgvspdData(lgCurrGrid).ReDraw = False
					
		If .MaxRows = 0 Then	' ù InsertRow�� 1��+�հ��� 

			iRow = 1			

			ggoSpread.InsertRow , 3
			Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
			
			.Row = 1
			.Col = C_W_TYPE : .Text = lgCurrGrid	
			.Col = C_SEQ_NO : .Text = 1	

			.Row = 2
			.Col = C_W_TYPE : .Text = lgCurrGrid	
			.Col = C_SEQ_NO : .Text = 1	
						
			.Row = 3
			.Col = C_W_TYPE : .Text = lgCurrGrid	
			.Col = C_SEQ_NO	: .Text = SUM_SEQ_NO	
			
			Call AddCellSpanRow(lgCurrGrid, 1)
			
			Call ReDrawTotalLine(lgCurrGrid)
			
			'ggoSpread.SpreadLock C_W9, iRow, .MaxCols-1, iRow
			Call SetSpreadLock

		Else
				
			If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
				ggoSpread.InsertRow iRow-1 , imRow
				SetSpreadColor lgCurrGrid, iRow, iRow + imRow - 1

				Call SetDefaultVal(lgCurrGrid, iRow, imRow)
			Else
				If (iRow Mod 2) = 1 Then
					iRow = iRow + 1
					.Row = iRow
				End If
				ggoSpread.InsertRow iRow ,imRow
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

Sub ReMakeGrid()
	Dim iRow, iMaxRows
	If lgvspdData(TYPE_2).MaxRows > 0 Then
		With lgvspdData(TYPE_2)
			iMaxRows = .MaxRows
			For iRow = 1 To iMaxRows Step 2
				.Row = iRow
				AddCellSpanRow TYPE_2, iRow
			Next
		End With
	ElseIf lgvspdData(TYPE_3).MaxRows > 0 Then
		With lgvspdData(TYPE_3)
			iMaxRows = .MaxRows
			For iRow = 1 To iMaxRows Step 2
				.Row = iRow
				AddCellSpanRow TYPE_3, iRow
			Next
		End With
	End If
End Sub

Sub AddCellSpanRow(Byval pType, Byval Row)
	Dim ret
	With lgvspdData(pType)
		.Row = Row
		ret = .AddCellSpan(C_W9		, Row, 1, 2)	
		ret = .AddCellSpan(C_W9_NM	, Row, 1, 2)
		ret = .AddCellSpan(C_W10	, Row, 1, 2)
		ret = .AddCellSpan(C_W10_NM	, Row, 1, 2)
		ret = .AddCellSpan(C_W11	, Row, 1, 2)
		ret = .AddCellSpan(C_W12	, Row, 1, 2)
		ret = .AddCellSpan(C_W13	, Row, 1, 2)
		ret = .AddCellSpan(C_W14	, Row, 1, 2)
		ret = .AddCellSpan(C_W15	, Row, 1, 2)
		ret = .AddCellSpan(C_W18_VIEW, Row, 1, 2)
		ret = .AddCellSpan(C_W18, Row, 1, 2)
		ret = .AddCellSpan(C_W18_VAL, Row, 1, 2)
		ret = .AddCellSpan(C_W19	, Row, 1, 2)
		ret = .AddCellSpan(C_W20	, Row, 1, 2)
		
		If pType = TYPE_2 Then
			ret = .AddCellSpan(C_W21	, Row, 2, 2)
		
		ElseIf pType = TYPE_3 Then
			ret = .AddCellSpan(C_W21	, Row, 1, 2)
			ret = .AddCellSpan(C_W35	, Row, 1, 2)
		End If
				
		.Col = C_W9_NM	: .TypeVAlign = 2
		.Col = C_W10_NM	: .TypeVAlign = 2
		.Col = C_W11	: .TypeVAlign = 2
		.Col = C_W12	: .TypeVAlign = 2
		.Col = C_W13	: .TypeVAlign = 2
		.Col = C_W14	: .TypeVAlign = 2
		.Col = C_W15	: .TypeVAlign = 2
		.Col = C_W18_VIEW: .TypeVAlign = 2
		.Col = C_W19	: .TypeVAlign = 2
		.Col = C_W20	: .TypeVAlign = 2
		.Col = C_W21	: .TypeVAlign = 2
		If pType = TYPE_3 Then
			.Col = C_W35	: .TypeVAlign = 2
		End If

	End With
End Sub

Sub RedrawTotalLine(Byval pType)
	Dim ret
	With lgvspdData(pType)
		.Row = .MaxRows
		.Col = C_W9_NM		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
		.Col = C_W10		: .CellType = 1
		.Col = C_W10_NM		: .CellType = 1
		.Col = C_W18_VIEW	: .CellType = 1	
	
		ret = .AddCellSpan(C_W10	, .MaxRows, 3, 1)
		ret = .AddCellSpan(C_W16	, .MaxRows, 2, 1)
		ret = .AddCellSpan(C_W20	, 1, 3, .MaxRows-1)
		If pType = TYPE_2 Then
			ret = .AddCellSpan(C_W21	, .MaxRows, 2, 1)
		End If
	End With
End Sub

' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(Index, iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 

	ggoSpread.Source = lgvspdData(lgCurrGrid)
	
	If iAddRows = 2 Then ' 1�ٸ� �ִ°�� 
		iSeqNo = MaxSpreadVal(lgvspdData(lgCurrGrid), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
		
		.Row = iRow
		.Col = C_W_TYPE : .Text = lgCurrGrid	
		.Col = C_SEQ_NO : .Text = iSeqNo
			
		.Row = iRow+1
		.Col = C_W_TYPE : .Text = lgCurrGrid	
		.Col = C_SEQ_NO : .Text = iSeqNo
		
		Call AddCellSpanRow(Index, iRow)
		
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(lgCurrGrid), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_W_TYPE : .Text = lgCurrGrid	
			.Col = C_SEQ_NO : .Text = iSeqNo
				
			.Row = i+1
			.Col = C_W_TYPE : .Text = lgCurrGrid	
			.Col = C_SEQ_NO : .Text = iSeqNo
		
			iSeqNo = iSeqNo + 1
			
			Call AddCellSpanRow(Index,i)
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows, iMaxRows, iRow, blnAllDel, iSeqNo

	blnAllDel = True

	With lgvspdData(lgCurrGrid) 
	
		If .MaxRows = 0 Then Exit Function
		
		ggoSpread.Source = lgvspdData(lgCurrGrid) 
		lDelRows = ggoSpread.DeleteRow

		If (.ActiveRow Mod 2) = 0 Then
			lDelRows = ggoSpread.DeleteRow( .ActiveRow-1)
		Else
			lDelRows = ggoSpread.DeleteRow( .ActiveRow+1)
		End If
		
		For iRow = 1 To .MaxRows
			.Row = iRow
			.Col = C_SEQ_NO : iSeqNo = UNICDbl(.Value)
			.Col = 0	
			If .Text <> ggoSpread.DeleteFlag And iSeqNo <> 999999 Then  blnAllDel = False ' -- ���� �ƴ��� & �谡 �ƴ����� ���� 
		Next
		
		If blnAllDel Then
			lDelRows = ggoSpread.DeleteRow(.MaxRows)
		Else
			ggoSpread.UpdateRow .MaxRows
		End If
			
		Call vspdData_Change(lgCurrGrid, C_W13, .ActiveRow)
		Call vspdData_Change(lgCurrGrid, C_W14, .ActiveRow)
		Call vspdData_Change(lgCurrGrid, C_W20, .ActiveRow)
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
        'strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows            
		
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
	
	If lgvspdData(TYPE_2).MaxRows > 0 Or lgvspdData(TYPE_3).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		Call ReMakeGrid
		
		If lgvspdData(TYPE_2).MaxRows > 0 Then
			Call ReDrawTotalLine(TYPE_2)
		ElseIf lgvspdData(TYPE_3).MaxRows > 0 Then
			Call ReDrawTotalLine(TYPE_3)
		End If
		
		Call SetSpreadLock()
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1101111100000111")										<%'��ư ���� ���� %>
			
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100000000000111")										<%'��ư ���� ���� %>
		End If
	Else
		Call SetToolbar("1100110100000111")										<%'��ư ���� ���� %>
	End If
			
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
    Dim lRow, lCol, lMaxRows, lMaxCols , i , sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    With frm1
		For i = C_W1 To C_W8
			strVal = strVal & .txtData(i).value &  Parent.gColSep
		Next
    End With
    frm1.txtSpread0.value = strVal
    strVal = ""
    
    For i = TYPE_2 To TYPE_3	' ��ü �׸��� ���� 
    
		With lgvspdData(i)
	
			ggoSpread.Source = lgvspdData(i)
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
			' ----- 1��° �׸��� 
			For lRow = 1 To lMaxRows Step 2
    
		       .Row = lRow
		       .Col = 0	: sTmp = Parent.gColSep

			  ' ��� �׸��� ����Ÿ ����     
			  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
					For lCol = 1 To lMaxCols
						.Col = lCol : sTmp = sTmp & Trim(.Text) &  Parent.gColSep
					Next
					
					If lRow <> lMaxRows Then
						.Row = lRow+1
						.Col = C_W16 : sTmp = sTmp & Trim(.Text) &  Parent.gColSep
						.Col = C_W17 : sTmp = sTmp & Trim(.Text) &  Parent.gColSep
					Else
						sTmp = sTmp & Parent.gColSep & Parent.gColSep
					End If
					sTmp = sTmp & Trim(.Text) &  Parent.gRowSep
			  End If  

		       .Col = 0
		       Select Case .Text
		           Case  ggoSpread.InsertFlag                                      '��: Insert
		                                              strVal = strVal & "C"  &  sTmp
		           Case  ggoSpread.UpdateFlag                                      '��: Update
		                                              strVal = strVal & "U"  &  sTmp
		           Case  ggoSpread.DeleteFlag                                      '��: Delete
		                                              strDel = strDel & "D"  &  sTmp
		       End Select
		       
			Next
		
		End With

		document.all("txtSpread" & CStr(i)).value =  strDel & strVal
		strDel = "" : strVal = ""
	Next
      
	frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
	frm1.txtFlgMode.value     = lgIntFlgMode
		

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
	Call InitVariables
	lgvspdData(TYPE_2).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_2)
    ggoSpread.ClearSpreadData

	lgvspdData(TYPE_3).MaxRows = 0
    ggoSpread.Source = lgvspdData(TYPE_3)
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
<SCRIPT LANGUAGE=javascript FOR=txtData EVENT=Change>
<!--
    SetHeadReCalc();
//-->
</SCRIPT>
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
						<a href="vbscript:GetRef">�ݾ� �ҷ�����</A>  
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
                                <TD WIDTH="100%"> 1. ���鼼�� ���</TD>
                            </TR>
                            <TR HEIGHT=10>
								<TD WIDTH="100%">
                                 <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
								     <TR>
								         <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(1)����</TD>
								         <TD CLASS="TD51" width="40%" ALIGN=CENTER COLSPAN=3>����ǥ�رݾ�</TD>
								         <TD CLASS="TD51" width="13%" ALIGN=CENTER ROWSPAN=2>(5)���⼼��</TD>
								         <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(6)������<BR>����ҵ�<BR>����[(2)/(4)]</TD>
								         <TD CLASS="TD51" width="10%" ALIGN=CENTER ROWSPAN=2>(7)�������<BR>(21), (35)</TD>
								         <TD CLASS="TD51" width="13%" ALIGN=CENTER ROWSPAN=2>(8)���� ����<BR>[(5) * (6) * (7)]</TD>
								    </TR>
								    <TR>
										<TD CLASS="TD51" width="15%" ALIGN=CENTER>(2)��������</TD>
										<TD CLASS="TD51" width="13%" ALIGN=CENTER>(3)�񰨸�����</TD>
										<TD CLASS="TD51" width="13%" ALIGN=CENTER>(4)��</TD>
								    </TR>
								    <TR>
								  		<TD><INPUT NAME="txtData" MAXLENGTH=25 TYPE="Text" ALT="(1)����" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="25"></TD>
								  		<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 85%></OBJECT>');</SCRIPT>
								  		<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenW2()"></TD>
								  		<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="25X2Z" width = 100%></OBJECT>');</SCRIPT></TD>
								  		<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT></TD>
								  		<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT></TD>
								  		<TD><INPUT NAME="txtData" MAXLENGTH=10 TYPE="Text" CLASS=FPDS140 ALT="" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%; text-align: 'center'"  tag="24X2"></TD>
								  		<TD><INPUT NAME="txtData" MAXLENGTH=10 TYPE="Text" CLASS=FPDS140 ALT="" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%; text-align: 'center'"  tag="24X2"></TD>
								  		<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtData" name=txtData CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2" width = 100%></OBJECT>');</SCRIPT></TD>
								    </TR>
								</table>
								</TD>
							</TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">2. �Ϲ����� ������� ���</TD>
                            </TR>
                            <TR HEIGHT=40%>
                                <TD WIDTH="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData0 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
                            </TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">3. �񰨸����� �����ϴ� �ܱ������ڱ�� ���� �Ǵ� �պ��ϴ� ����� ������� ���</TD>
                            </TR>   
                            <TR HEIGHT=40%>
                                <TD WIDTH="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>�Ϲ����� �������</LABEL>&nbsp;
				           <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>�񰨸����� �����ϴ� �ܱ������ڱ�� ���� �Ǵ� �պ��ϴ� ����� �������</LABEL>&nbsp;
				           
				        </TD>
				            
                </TR>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FNCBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24" style="display='none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" style="display='none'"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" style="display='none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

