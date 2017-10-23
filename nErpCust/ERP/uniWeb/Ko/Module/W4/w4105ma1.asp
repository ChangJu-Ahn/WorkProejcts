
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �غ������ 
'*  3. Program ID           : W4105MA1
'*  4. Program Name         : W4105MA1.asp
'*  5. Program Desc         : ��5ȣ Ư����������� 
'*  6. Modified date(First) : 2005/01/18
'*  7. Modified date(Last)  : 2005/01/18
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

Const BIZ_MNU_ID		= "W4105MA1"
Const BIZ_PGM_ID		= "W4105mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "W4105mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID	    = "W4105OA1"

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 

' -- �׸��� �÷� ���� 
Dim C_SEQ_NO
Dim C_W1_CD
Dim C_W1
Dim C_W2
Dim C_W2_CD
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(2)

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO	= 1
	C_W1_CD		= 2	' -- 1�� �׸��� 
	C_W1		= 3	' ���� 
	C_W2		= 4	' �ٰŹ����� 
	C_W2_CD		= 5 ' �ڵ� 
	C_W3		= 6	' ȸ����� 
	C_W4		= 7 ' �ѵ��ʰ��� 
	C_W5		= 8	' ������ 
	C_W6		= 9	' �����Ѽ�����ձݺ��ξ� 
	C_W7		= 10	' �ձݺһ��Ծװ� 
	C_W8		= 11	' �ձݻ��԰� 
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
    
End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	
    Call initSpreadPosVariables()  

	' 1�� �׸��� 
	With lgvspdData(TYPE_1)
			
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W8 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		'����� 3�ٷ�    
		.ColHeaderRows = 2    
						       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'Call AppendNumberPlace("6","4","0")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1_CD,	"MINOR_CD", 10,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit     C_W1,	    "(1)�� ��", 40,,,60,1 
		ggoSpread.SSSetEdit		C_W2,		"�ٰŹ�����"	, 12,,,50,1 
		ggoSpread.SSSetEdit		C_W2_CD,	"�ڵ�", 5,2,,2,1 
		ggoSpread.SSSetFloat	C_W3,		"(2)ȸ�����"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W4,		"(3)�ѵ��ʰ���"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W5,		"(4)������"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
		ggoSpread.SSSetFloat	C_W6,		"(5)�����Ѽ�" & vbCrLf & "����" & vbCrLf & "�ձݺ��ξ�"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_W7,		"(6)�ձݺһ��԰�" & vbCrLf & "[(3)+(5)]" , 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W8,		"(7)�ձݻ��԰�" & vbCrLf & "[(2)-(6)]"	, 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	

		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SEQ_NO , -1000, 1, 2)
		ret = .AddCellSpan(C_W1_CD	, -1000, 1, 2)	' ���� 2�� ��ħ 
		ret = .AddCellSpan(C_W1		, -1000, 1, 2)	
		ret = .AddCellSpan(C_W2		, -1000, 1, 2)
		ret = .AddCellSpan(C_W2_CD	, -1000, 1, 2)
		ret = .AddCellSpan(C_W3		, -1000, 3, 1)
		ret = .AddCellSpan(C_W6		, -1000, 1, 2)
		ret = .AddCellSpan(C_W7		, -1000, 1, 2)
		ret = .AddCellSpan(C_W8		, -1000, 1, 2)
    
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W3	: .Text = "�� �� �� �� �� �� ��"
		
		.Row = -999
		.Col = C_W3	: .Text = "(2)ȸ�����"
		.Col = C_W4	: .Text = "(3)�ѵ��ʰ���"
		.Col = C_W5	: .Text = "(4)������"
			
		.rowheight(-999) = 20	' ���� ������	(2���� ���, 1���� 15)
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W1_CD,C_W1_CD,True)
		Call ggoSpread.SSSetColHidden(C_W2,C_W2,True)
				
		Call SetSpreadLock(TYPE_1)
				
		.ReDraw = true	
			
	End With 
  
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()

	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
       
    ' �׸��� �ʱ� ����Ÿ���� 
    Dim arrMinorCd, arrW1, arrW2, arrW2_CD, iMaxRows, iRow, iMinorCnt
	'call CommonQueryRs("MINOR_CD, MINOR_NM, REFERENCE_1, REFERENCE_2","ufn_TB_Configuration('W1054', '" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
    'arrMinorCd	= Split(lgF0, Chr(11))
    'arrW1		= Split(lgF1, Chr(11))
    'arrW2		= Split(lgF2, Chr(11))
    'arrW2_CD	= Split(lgF3, Chr(11))
    
    'iMinorCnt = 18	' ���̳��ڵ� �� ������ 18���� �ϵ��ڵ� 
	'iMaxRows = UBound(arrMinorCd)
	iMaxRows = 22
	
	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		ggoSpread.InsertRow , iMaxRows
		
		' �迭�� �׸��忡 ���� 
		For iRow = 1 To iMaxRows
			
			.Row = iRow
			.Col = C_SEQ_NO : .value = iRow
			.Col = C_W1_CD	: .value = 100 + iRow
		'	.Col = C_W1		: .value = arrW1(iRow-1)
		'	.Col = C_W2		: .value = arrW2(iRow-1)
		'	.Col = C_W2_CD	: .value = arrW2_CD(iRow-1)
			
			
		'	If Trim(arrW1(iRow-1)) <> "" Then
		'		ggoSpread.SpreadLock C_W1, iRow, C_W2_CD, iRow
		'		ggoSpread.SpreadLock C_W5, iRow, C_W5, iRow
		'		ggoSpread.SpreadLock C_W7, iRow, C_W8, iRow
		'	End If
			
		Next
		
		'If iMinorCnt > iMaxRows Then
		'	ggoSpread.InsertRow iMaxRows, iMinorCnt - iMaxRows + 4
		'	
		'	For iRow = iMaxRows To iMinorCnt + 4
		'		.Row = iRow
		'		.Col = C_SEQ_NO : .value = iRow			
		'	Next
		'End If

		'ggoSpread.SpreadLock C_W1, 1, C_W2_CD, 8
		'ggoSpread.SpreadLock C_W5, iRow, C_W5, iRow
		'ggoSpread.SpreadLock C_W7, iRow, C_W8, iRow
				
		.Row = 1
		.Col = C_W1		: .Value = "(101)�ֱǻ����߼ұ�� ���� ����ս��غ��(��8����2)"
		.Col = C_W2_CD	: .Value = "42"
		
		.Row = 2
		.Col = C_W1		: .Value = "(102)�������η°����غ��(��9��)"
		.Col = C_W2_CD	: .Value = "02"

		.Row = 3
		.Col = C_W1		: .Value = "(103)��ȸ�����ں������غ��(��28��)"
		.Col = C_W2_CD	: .Value = "08"

		.Row = 4
		.Col = C_W1		: .Value = "(104)�ε�������ȸ�����ڼս��غ��(��55����2)"
		.Col = C_W2_CD	: .Value = "45"

		.Row = 5
		.Col = C_W1		: .Value = "(105)100%�ձݻ��԰������� ����غ��(��74����1��)"
		.Col = C_W2_CD	: .Value = "17"

		.Row = 6
		.Col = C_W1		: .Value = "(106)80%�ձݻ��԰������� ����غ��(��74����2��)"
		.Col = C_W2_CD	: .Value = "18"

		.Row = 7
		.Col = C_W1		: .Value = "(107)�ֱǻ����� ���� �ڻ��� ó�мս��غ��(��104��3)"
		.Col = C_W2_CD	: .Value = "43"

		.Row = 8
		.Col = C_W1		: .Value = "(108)��ȭ����غ��(��104��9)"
		.Col = C_W2_CD	: .Value = "47"

		.Row = 9
		.Col = C_W1		: .Value = "(109)"
		.Col = C_W2_CD	: .Value = "44"

		.Row = 10
		.Col = C_W1		: .Value = "(110)"
		.Col = C_W2_CD	: .Value = ""

		.Row = 11
		.Col = C_W1		: .Value = "(111)"
		.Col = C_W2_CD	: .Value = ""

		.Row = 12
		.Col = C_W1		: .Value = "(112)"
		.Col = C_W2_CD	: .Value = ""
		
		.Row = 13
		.Col = C_W1		: .Value = "(113)"

		.Row = 14
		.Col = C_W1		: .Value = "(114)"

		.Row = 15
		.Col = C_W1		: .Value = "(115)"

		.Row = 16
		.Col = C_W1		: .Value = "(116)"
		
		.Row = 17
		.Col = C_W1		: .Value = "(117)"

		.Row = 18
		.Col = C_W1		: .Value = "(118)"

		
		
		' �غ�ݰ�(119), �غ�� �� Ư�������󰢺� ��(142)�� ȸ�� ó�� 
		.Row = 19
		'.Col = C_W1_CD	: .Value = "119"
		.Col = C_W1		: .Value = "(119)�غ�� ��"
		.Col = C_W2_CD	: .Value = "19"

		.Row = 20
		'.Col = C_W1_CD	: .Value = "140"
		.Col = C_W1		: .Value = "(140)Ư�������󰢺� ��"
		.Col = C_W2_CD	: .Value = "40"

		.Row = 21
		'.Col = C_W1_CD	: .Value = "141"
		.Col = C_W1		: .Value = "(141)Ư���ڻ갨���󰢺� ��(��30��)"
		.Col = C_W2_CD	: .Value = "46"

		.Row = 22
		'.Col = C_W1_CD	: .Value = "142"
		.Col = C_W1		: .Value = "�غ�� �� Ư�������󰢺� ��(119+140+141)"
		.Col = C_W2_CD	: .Value = "41"

		Call SetSpreadLock_Query(TYPE_1)
				
		lgvspdData(TYPE_1).Redraw = True
	End With

End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	Select Case pType
		Case TYPE_1
			ggoSpread.SpreadLock C_SEQ_NO, -1, C_W1_CD
			ggoSpread.SpreadLock C_W5, -1, C_W5
			ggoSpread.SpreadLock C_W7, -1, C_W7
			ggoSpread.SpreadLock C_W8, -1, C_W8
	End Select
	
End Sub

Sub SetSpreadLock_Query(Byval pType)
	Dim iRow

	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		ggoSpread.SpreadLock C_SEQ_NO, -1, C_W1_CD
		ggoSpread.SpreadLock C_W1, 1, C_W1, 8
		ggoSpread.SpreadLock C_W2_CD, 1, C_W2_CD, 9
		ggoSpread.SpreadLock C_W2_CD,19, C_W2_CD, 22
		ggoSpread.SpreadLock C_W5, -1, C_W5
		ggoSpread.SpreadLock C_W7, -1, C_W7
		ggoSpread.SpreadLock C_W8, -1, C_W8
		ggoSpread.SpreadLock C_SEQ_NO, 20, C_W2_CD, 21
		ggoSpread.SpreadLock C_SEQ_NO, 19, C_W8, 19
		ggoSpread.SpreadLock C_SEQ_NO, 22, C_W8, 22
				
		lgvspdData(TYPE_1).Redraw = True
	End With
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
	ggoSpread.Source = lgvspdData(pType)

	Select Case pType
		Case TYPE_1
	End Select
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"

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
			
    ' �׸��� �ʱ� ����Ÿ���� 
	call CommonQueryRs("W5, W6","TB_31_1H "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		.Row = 1
		.Col = C_W3 : .Value = Replace(lgF0, chr(11), "")
		.Col = C_W4 : .Value = Replace(lgF1, chr(11), "")
		
		Call vspdData_Change(TYPE_1, C_W3, 1)
		Call vspdData_Change(TYPE_1, C_W4, 1)
		
		lgvspdData(TYPE_1).Redraw = True
	End With

	call CommonQueryRs("W4, W5","TB_31_2H "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	With lgvspdData(TYPE_1)
		lgvspdData(TYPE_1).Redraw = False
		
		ggoSpread.Source = lgvspdData(TYPE_1)
		
		.Row = 3
		.Col = C_W3 : .Value = Replace(lgF0, chr(11), "")
		.Col = C_W4 : .Value = Replace(lgF1, chr(11), "")
		
		Call vspdData_Change(TYPE_1, C_W3, 3)
		Call vspdData_Change(TYPE_1, C_W4, 3)
		
		lgvspdData(TYPE_1).Redraw = True
	End With
	
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
'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000101111")										<%'��ư ���� ���� %>
	  
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
	'Call GetFISC_DATE
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
	Dim dblSum, dblAmt(20), dblW3, dblW4, dblW5, dblW6, dblW7, dblW8
	Dim sCoCd,sFiscYear,sRepType
	
	sCoCd		= "<%=wgCO_CD%>"
	sFiscYear	= "<%=wgFISC_YEAR%>"
	sRepType	= "<%=wgREP_TYPE%>"
	' �������� ��� 
	
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

	lgvspdData(Index).Redraw = False
	
	If Index = TYPE_1 Then	'1�� �׸� 
	
		Select Case Col
			Case C_W3, C_W4, C_W6	' ȸ�����/�ѵ��ʰ��� 
			
				.Col = C_W3
				
				IF .Row = 1 THEN 
					Call CommonQueryRs("W5"," TB_31_1H(NOLOCK) ","  CO_CD = '" & sCoCd & "' and FISC_YEAR = '" & sFiscYear & "' and REP_TYPE =  '" & sRepType & "' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					if UNICDbl(.Value) <> UNICDbl(Replace(lgF0,Chr(11),""))  then
						Call DisplayMsgBox("WC0004", parent.VB_INFORMATION, "Ư����� ��������", "�߼ұ�������غ�� ��������")
					   .Value = UNICDbl(Replace(lgF0,Chr(11),""))
					End if 
				End IF
				.Row = Row 
				.Col = C_W3 :dblAmt(C_W3) = UNICDbl(.Value)
				.Col = C_W4 :dblAmt(C_W4) = UNICDbl(.Value)
				.Col = C_W6 :dblAmt(C_W6) = UNICDbl(.Value)
				
				dblAmt(C_W5) = dblAmt(C_W3) - dblAmt(C_W4)
				If dblAmt(C_W5) < 0 Then
					Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, GetGrid(C_W4, -999), GetGrid(C_W3, -999))
					.Row = Row
					.Col = Col	: .value = 0
					.Col = Col	: dblAmt(Col) = 0
					dblAmt(C_W5) = dblAmt(C_W3) - dblAmt(C_W4)
				End If

				dblAmt(C_W7) = dblAmt(C_W4) + dblAmt(C_W6)	
				dblAmt(C_W8) = dblAmt(C_W3) - dblAmt(C_W7)
				If dblAmt(C_W8) < 0 Then
					Call DisplayMsgBox("WC0010", parent.VB_INFORMATION, GetGrid(C_W7, 0), GetGrid(C_W3, -999))
					.Row = Row
					.Col = Col	: .value = 0
					.Col = Col	: dblAmt(Col) = 0
					dblAmt(C_W7) = dblAmt(C_W4) + dblAmt(C_W6)
					dblAmt(C_W8) = dblAmt(C_W3) - dblAmt(C_W7)
				End If
						
				.Col = C_W5 : .Value = dblAmt(C_W5)
				.Col = C_W7 : .Value = dblAmt(C_W7)
				.Col = C_W8 : .Value = dblAmt(C_W8)

				' -- ���� ����� �÷��� �غ�� ���� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, 19 - 1, true, 19, Col, "V")	' �հ� 
				
				' -- C_W5 ������ �غ�� ���� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W5, 1, 19 - 1, true, 19, C_W5, "V")	' �հ� 
				' -- C_W7
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W7, 1, 19 - 1, true, 19, C_W7, "V")	' �հ� 
				' -- C_W8
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W8, 1, 19 - 1, true, 19, C_W8, "V")	' �հ� 
				
				ggoSpread.UpdateRow 19

				
				' -- ���� ����� �÷��� �غ�� ���� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 19, .MaxRows- 1, true, .MaxRows, Col, "V")	' �հ� 
				
				' -- C_W5 ������ �غ�� ���� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W5, 19, .MaxRows- 1 - 1, true, .MaxRows, C_W5, "V")	' �հ� 
				' -- C_W7
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W7, 19, .MaxRows- 1 - 1, true, .MaxRows, C_W7, "V")	' �հ� 
				' -- C_W8
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W8, 19, .MaxRows- 1 - 1, true, .MaxRows, C_W8, "V")	' �հ�				
				
				ggoSpread.UpdateRow .MaxRows
		End Select
	
	End If
	
	lgvspdData(Index).Redraw = True
	
	End With
	
End Sub

Function GetGrid(Byval Col, Byval Row)
	With lgvspdData(TYPE_1)
		.Col = Col : .Row = Row : GetGrid = .value
	End With
End Function

' -- 2��° �׸��� 
Sub SetGridTYPE_2()
	Dim dblSum, dblW11, dblW12, dblW13 , dblFiscYear, dblW26, dblW25, dblW24, dblW23, dblW22, dblW17, dblW15, dblW14
	Dim dblW27, dblW28, dblW29, dblW30, dblW31, dblW32, dblW33

	With lgvspdData(TYPE_2)
		.Row = .ActiveRow
		.Col = C_W19 : dblW19 = UNICDbl(.Value)
		.Col = C_W20 : dblW20 = UNICDbl(.Value)
		.Col = C_W21 : dblW21 = UNICDbl(.Value)
									
		' �հ躯�� 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W19, 1, .MaxRows - 1, true, .MaxRows, C_W19, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W20, 1, .MaxRows - 1, true, .MaxRows, C_W20, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W21, 1, .MaxRows - 1, true, .MaxRows, C_W21, "V")	' �հ� 
					
		' W22 ���� 
		dblW22 = dblW19 + dblW20 + dblW21
		.Col = C_W22	: .Row = .ActiveRow : .Value = dblW22
					
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W22, 1, .MaxRows - 1, true, .MaxRows, C_W22, "V")	' �հ� 
					
		' W23 ���� 
		.Col = C_W17	: .Row = .ActiveRow : dblW17 = UNICDbl(.value)
		dblW23 = dblW17 + dblW22
		.Col = C_W23	: .Row = .ActiveRow : .Value = dblW23

		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W23, 1, .MaxRows - 1, true, .MaxRows, C_W23, "V")	' �հ� 
		
		.Row = .ActiveRow			
		.Col = C_W24	: dblW24 = UNICDbl(.Value)
		' W25 ���� 
		dblW25= dblW23 - dblW24
		.Col = C_W25	: .Value = dblW25
					
		Call FncSumSheet(lgvspdData(lgCurrGrid), C_W25, 1, .MaxRows - 1, true, .MaxRows, C_W25, "V")	' �հ�	
	End With
End Sub

' 2�� �׸��忡�� 1�� �׸����� ����Ÿ�� ã�Ƽ� W16�ݾ��� �����Ѵ� 
Sub GetW16(Byval pYear , Byref pdblW16, Byref pdblW17)
	Dim iRow, iMaxRows
	With lgvspdData(TYPE_2)
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


' ��� ���� 
Sub SetHeadReCalc()	
	Dim dblSum, dblW16, dblW3, dblW2, dblW4, dblW5, dblW6, dblW7
	
	With lgvspdData(TYPE_1)
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
End Sub

' �ܾ� �÷��� ����ɶ� ȣ��� 
Sub SetW17_W18(Row)
	Dim dblSum, dblW4, dblW15, dblW16, dblW18
	
	With lgvspdData(TYPE_1)
		
		.Row = Row
		.Col = C_W11	: dblW11 = UNICDbl(.Value)	' ���� 
		.Col = C_W12	: dblW12 = UNICDbl(.Value)	' ���� 
		.Col = C_W13	: dblW13 = UNICDbl(.Value)	' ���� 
		
		.Col = C_W14	: dblW14 = UNICDbl(.Value)	' ���� 
		.Col = C_W15	: dblW15 = UNICDbl(.Value)	' �뺯 
		.Col = C_W16	: dblW16 = UNICDbl(.Value)	' ���� 
		.Col = C_W18	: dblW18 = UNICDbl(.Value)	' �뺯 
		
		.Col = C_W18	: dblW18 = dblW11 - dblW12 - dblW13				: .Value = dblW18
		.Col = C_W17	: dblW17 = dblW18 - dblW14 - dblW15 - dblW16	: .Value = dblW17

		Call FncSumSheet(lgvspdData(TYPE_2), C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' �հ� 
		Call FncSumSheet(lgvspdData(TYPE_2), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' �հ� 
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
    '   ggoSpread.Source = lgvspdData(Index)
    '   
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
    'Call InitData                              
    															
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
	ggoSpread.Source = lgvspdData(TYPE_1)
	If ggoSpread.SSCheckChange = True Then
		blnChange = True
	End If
	
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

' ---------------------- ���ĳ� ���� -------------------------
Function  Verification()

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

    Call SetToolbar("1100100000001111")

	'Call ClickTab1()
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
    
    Call CheckReCalc()				' �Ѷ����� ��ҵǸ� ���� 
    Call CheckW7Status(lgCurrGrid)	' ������ ���� üũ 
End Function

' ���� 
Function CheckReCalc()
	Dim dblSum
	
	With lgvspdData(lgCurrGrid)
		ggoSpread.Source = lgvspdData(lgCurrGrid)	
	
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' �հ� 
		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' �հ� 
	
		' �ڽ��� ����Ǹ� �ٸ� ȣ���ؾ� �� �κ��� �Ʒ��� ��� 
		Call SetW5(lgCurrGrid, 1)
					
		' �ϼ� ���� ��û 
		Call SetW6(lgCurrGrid, 1)
					
		Call SetW7(lgCurrGrid, 1)	' ���� ��� 
	End With
End Function

Function FncInsertRow(ByVal pvRowCnt) 
  
    
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
	End With
	
	Call CheckReCalc()				' �Ѷ����� ��ҵǸ� ���� 
	
	Call CheckW7Status(lgCurrGrid)	' ������ ���� üũ 
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


Function DbQueryFalse()
	Call InitData
End Function


Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If lgvspdData(TYPE_1).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
			Call SetSpreadLock_Query(TYPE_1)
			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1101100000011111")										<%'��ư ���� ���� %>
			
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100100000011111")										<%'��ư ���� ���� %>
		End If
	Else
		Call SetToolbar("1100100000111111")										<%'��ư ���� ���� %>
	End If
	
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
    
    For i = TYPE_1 To TYPE_1	' ��ü �׸��� ���� 
    
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
									<TD CLASS="TD6"><script language =javascript src='./js/w4105ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>

						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
							<TR>
							     <TD WIDTH="100%" VALIGN=TOP HEIGHT=* VALIGN=TOP><script language =javascript src='./js/w4105ma1_vspdData0_vspdData0.js'></script></TD>
									  </TR>
							
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
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

