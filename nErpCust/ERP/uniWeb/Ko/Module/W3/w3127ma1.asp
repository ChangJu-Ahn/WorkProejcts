
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �� ���� ���� 
'*  3. Program ID           : W3127MA1
'*  4. Program Name         : W3127MA1.asp
'*  5. Program Desc         : ��26ȣ ���������ε��� ������ ���Ա�������������(��)
'*  6. Modified date(First) : 2005/01/05
'*  7. Modified date(Last)  : 2006/02/08
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : HJO 
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

Const BIZ_MNU_ID		= "W3127MA1"
Const BIZ_PGM_ID		= "w3127mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "w3127mb2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "W3127OA1"

Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2
Const TAB3 = 3

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.
Const TYPE_3	= 2		
Const TYPE_4A	= 3
Const TYPE_4B	= 4
Const TYPE_5	= 5
Const TYPE_6	= 6

' -- �׸��� �÷� ���� 
Dim C_SEQ_NO
Dim C_W_TYPE
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7

Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13
Dim C_W14

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(6)
Dim lgFISC_START_DT, lgFISC_END_DT, lgW12_REF ' �ں���+������.....

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	
	C_SEQ_NO	= 1	' -- 1�� �׸��� 
    C_W_TYPE	= 2	' ���� 
    C_W1		= 3	' ������ 
    C_W2		= 4 ' ���� 
    C_W3		= 5	' ���� 
    C_W4		= 6	' �뺯 
    C_W5		= 7	' �ܾ� 
    C_W6		= 8	' �ϼ� 
    C_W7		= 9	' ���� 

	C_W10		= 5	' �ڻ��Ѱ� 
	C_W11		= 6	' ��ä�Ѱ� 
	C_W12		= 7	' �ڱ��ں� 
	C_W13		= 8 ' ���ϼ� 
	C_W14		= 9	' ���� 
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
    
    lgW12_REF = 0
    lgCurrGrid = TYPE_1
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
	Dim ret, iRow
	
	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1
	Set lgvspdData(TYPE_3) = frm1.vspdData2
	Set lgvspdData(TYPE_4A) = frm1.vspdData3
	Set lgvspdData(TYPE_4B) = frm1.vspdData4
	Set lgvspdData(TYPE_5) = frm1.vspdData5
	Set lgvspdData(TYPE_6) = frm1.vspdData6
	
	lgvspdData(TYPE_1).ScriptEnhanced  = True
	lgvspdData(TYPE_2).ScriptEnhanced  = True
	lgvspdData(TYPE_3).ScriptEnhanced  = True
	lgvspdData(TYPE_4A).ScriptEnhanced  = True
	lgvspdData(TYPE_4B).ScriptEnhanced  = True
	lgvspdData(TYPE_5).ScriptEnhanced  = True
	lgvspdData(TYPE_6).ScriptEnhanced  = True
	
    Call initSpreadPosVariables()  

	' 1��-5�� �׸��� 

	For iRow = TYPE_1 To TYPE_5		' �׸��� ���� ��� 
		With lgvspdData(iRow)
			
			ggoSpread.Source = lgvspdData(iRow)	
			'patch version
			ggoSpread.Spreadinit "V20041222_" & iRow,,parent.gForbidDragDropSpread    
    
			.ReDraw = false

			.MaxCols = C_W7 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.Col = .MaxCols									'��: ����� �� Hidden Column
			.ColHidden = True    
				       
			.MaxRows = 0
			ggoSpread.ClearSpreadData

			'Call AppendNumberPlace("6","3","2")

			ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
			ggoSpread.SSSetEdit		C_W_TYPE,	"�� ��"		, 10,,,6,1	' �����÷� 
			ggoSpread.SSSetDate		C_W1,		"(1)������"	, 10, 2, Parent.gDateFormat, -1
			ggoSpread.SSSetEdit		C_W2,		"(2)�� ��"	, 15,,,50,1
			ggoSpread.SSSetFloat	C_W3,		"(3)�� ��"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W4,		"(4)�� ��"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W5,		"(5)�� ��"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			ggoSpread.SSSetFloat	C_W6,		"(6)�ϼ�"	, 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	
			ggoSpread.SSSetFloat	C_W7,		"(7)�� ��"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
			Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
				
			Call SetSpreadLock(iRow)
				
			.ReDraw = true	
			
		End With 
	Next
 
	With lgvspdData(TYPE_6)
	
		' �ڱ��ں� ������� �׸��� ���� 
 		ggoSpread.Source = lgvspdData(TYPE_6)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_6,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W14 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		ggoSpread.ClearSpreadData
		.MaxRows = 0
		'Call AppendNumberPlace("6","3","2")

		ggoSpread.SSSetEdit		C_SEQ_NO,	"����"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W_TYPE,	"�� ��"		, 10,,,6,1	' �����÷� 
		ggoSpread.SSSetDate		C_W1,		"Blank"			, 10, 2, Parent.gDateFormat, -1
		ggoSpread.SSSetFloat	C_W2,		"Blank"			, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W10,		"(10)��������ǥ" & vbCrLf & "�ڻ��Ѱ�", 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W11,		"(11)��������ǥ" & vbCrLf & "��ä�Ѱ�", 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W12,		"(12)�ڱ��ں� [(10)-(11)]"	, 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W13,		"(13)�������" & vbCrLf & "�ϼ�", 10, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	C_W14,		"(14)�� ��"	, 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		Call ggoSpread.SSSetColHidden(C_W_TYPE,C_W_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
		Call ggoSpread.SSSetColHidden(C_W2,C_W2,True)
		
		.rowheight(-1000) = 20	' ���� ������	(2���� ���, 1���� 15)
		
		Call SetSpreadLock(TYPE_6)
			
		.ReDraw = true	 
    
    End With
     
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
    
	With lgvspdData(TYPE_6)
		ggoSpread.Source = lgvspdData(TYPE_6)
		
		ggoSpread.InsertRow ,1
		SetSpreadColor TYPE_6, 1, 1
		
		.Col = C_SEQ_NO : .Row = 1 : .text = 1
		.Col = C_W_TYPE : .Row = 1 : .text = TYPE_6
	End With

	Call GetFISC_DATE

End Sub

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock(Byval pType)

	ggoSpread.Source = lgvspdData(pType)	
	
	If pType = TYPE_6 Then	' �ڱ��ں�������� 
		ggoSpread.SpreadLock C_W13, -1, C_W13
		ggoSpread.SpreadLock C_W14, -1, C_W14
	Else
		ggoSpread.SpreadLock C_W5, -1, C_W5
		ggoSpread.SpreadLock C_W6, -1, C_W6
		ggoSpread.SpreadLock C_W7, -1, C_W7
		ggoSpread.SSSetRequired	 C_W1, -1, C_W1
		'ggoSpread.SSSetProtected C_W1, lgvspdData(pType).MaxRows, C_W1, lgvspdData(pType).MaxRows 
	End If
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	ggoSpread.Source = lgvspdData(pType)

	If pType = TYPE_6 Then	' �ڱ��ں�������� 
		ggoSpread.SSSetProtected C_W13, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W14, pvStartRow, pvEndRow 
	Else
		ggoSpread.SSSetProtected C_W5, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W6, pvStartRow, pvEndRow 
		ggoSpread.SSSetProtected C_W7, pvStartRow, pvEndRow 
		If lgvspdData(pType).MaxRows = pvEndRow Then
			ggoSpread.SSSetRequired	 C_W1, pvStartRow, pvEndRow -1
		Else
			ggoSpread.SSSetRequired	 C_W1, pvStartRow, pvEndRow
		End If
	End If

End Sub

Sub SetSpreadTotalLine()
	Dim iRow
	For iRow = TYPE_1 To TYPE_5
		ggoSpread.Source = lgvspdData(iRow)
		With lgvspdData(iRow)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
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
            C_W21		= iCurColumnPos(3)
            C_W1		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W21		= iCurColumnPos(6)
            C_W3		= iCurColumnPos(7)
            C_W4		= iCurColumnPos(8)
            C_W5		= iCurColumnPos(9)
            C_W13		= iCurColumnPos(10)
            C_W15		= iCurColumnPos(11)
            C_W16		= iCurColumnPos(12)
            C_W17		= iCurColumnPos(13)
            C_W_TYPE	= iCurColumnPos(14)
            C_W1		= iCurColumnPos(15)
            C_W2		= iCurColumnPos(16)
    End Select    
End Sub

'============================== ���۷��� �Լ�  ========================================

Sub GetFISC_DATE()	' ��û������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd
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
	
	' 6. �ڱ��ں� ������� �׸��� W13 �Է� 
	With lgvspdData(TYPE_6)
		.Col = C_W13 : .Row = .MaxRows 
		If frm1.cboREP_TYPE.value = "2" Then
			.text = DateDiff("d", lgFISC_START_DT, DateAdd("m", 6, lgFISC_START_DT) - 1)+1
		Else
			.text = lgFISC_END_DT - lgFISC_START_DT + 1
		End If
	End With
End Sub

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

'    iCalledAspName = AskPRAspName("W5105RA1")
    
 '   If Trim(iCalledAspName) = "" Then
  '      IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "W5105RA1", "x")
   '     IsOpenPop = False
    '    Exit Function
    'End If
    
'    With frm1
 '       If .vspdData.ActiveRow > 0 then 
  '          arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
   '         arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
    '    End If            
    'End With    

    arrRet = window.showModalDialog("../W5/W5105RA1.asp", Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

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
	lgCurrGrid = TYPE_3
End Function

Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3)
	gSelframeFlg = TAB3
	lgCurrGrid = TYPE_5
End Function


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

	Call InitData 
	Call fncQuery
     
    
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
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_3
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 3�� �׸��� 
Sub vspdData3_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_4A
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4A
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_4A
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_4A
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData3_GotFocus()
	lgCurrGrid = TYPE_4A
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData3_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_4A
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_4A
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_4A
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 4�� �׸��� 
Sub vspdData4_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_4B
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_4B
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_4B
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData4_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_4B
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData4_GotFocus()
	lgCurrGrid = TYPE_4B
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData4_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_4B
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData4_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_4B
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData4_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_4B
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 5�� �׸��� 
Sub vspdData5_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_5
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData5_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_5
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData5_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_5
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData5_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_5
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData5_GotFocus()
	lgCurrGrid = TYPE_5
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData5_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_5
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData5_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_5
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData5_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_5
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

' -- 6�� �׸��� 
Sub vspdData6_Change(ByVal Col , ByVal Row )
	lgCurrGrid = TYPE_6
	Call vspdData_Change(lgCurrGrid, Col, Row)
End Sub

Sub vspdData6_Click(ByVal Col, ByVal Row)
	lgCurrGrid = TYPE_6
	Call vspdData_Click(lgCurrGrid,  Col,  Row)
End Sub

Sub vspdData6_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	lgCurrGrid = TYPE_6
	Call vspdData_ColWidthChange(lgCurrGrid, pvCol1, pvCol2)
End Sub

Sub vspdData6_DblClick(ByVal Col, ByVal Row)				
	lgCurrGrid = TYPE_6
	Call vspdData_DblClick(lgCurrGrid, Col, Row)
End Sub

Sub vspdData6_GotFocus()
	lgCurrGrid = TYPE_6
	Call vspdData_GotFocus(lgCurrGrid)
End Sub

Sub vspdData6_MouseDown(Button , Shift , x , y)
	lgCurrGrid = TYPE_6
	Call vspdData_MouseDown(lgCurrGrid, Button , Shift , x , y)
End Sub    

Sub vspdData6_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	lgCurrGrid = TYPE_6
	Call vspdData_ScriptDragDropBlock(lgCurrGrid, Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
End Sub 
Sub vspdData6_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	lgCurrGrid = TYPE_6
	Call vspdData_TopLeftChange(lgCurrGrid, OldLeft, OldTop, NewLeft, NewTop)
End Sub

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, datW1_DOWN, datW1, iRow, iMaxRows, dblW5_UP, dblW10, dblW11, dblW3, dblW4, dblW5
	Dim preSum
	
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
	
	If Index = TYPE_6 Then	' �ڱ��ں�������� 
		Select Case Col
			Case C_W10, C_W11
				.Col = C_W10 : dblW10 = UNICDbl(.text)
				.Col = C_W11 : dblW11 = UNICDbl(.text)
				
				If (dblW10 - dblW11) > lgW12_REF Then
					.Col = C_W12 : .text = dblW10 - dblW11
				Else
					.Col = C_W12 : .text = lgW12_REF
				End If
				
				Call SetW14()
			Case C_W12	' -- 2006.03.22 �̺�Ʈ �߰�
				Call SetW14()
		End Select 
		
	ElseIf Index=TYPE_4B Then	'�뺯������ �ٸ� �Ͱ� �ݴ�
		Select Case Col
			Case C_W1	' ������ ����� 
				If lgFISC_END_DT = "" Then
					MsgBox "���� �⺻������ ������� �������� ����ֽ��ϴ�"
					Exit Sub
				Else
					iMaxRows = .MaxRows
				
					' 1-1. ���� �Է��� �������� �������� �����ຸ�� ũ�� ������ ����Ų��.
					If Row + 1 <> iMaxRows Then
						.Row = Row		: .Col = C_W1	: datW1 = CDate(.Text)
						
						' 1.1 �Ʒ����� ���� ��� 
						.Row = Row+1	: .Col = C_W1	
						If .Text <> "" Then
							datW1_DOWN = CDate(.Text)

							If datW1 > datW1_DOWN Then ' �Ʒ��ຸ�� ��¥�� ���ĸ� ���� 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '��: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
					'1-2.�����Է��� �������� �������� �����ຸ�� ������ ����
					ElseIf Row-1 <> 0 Then 
						.Row=Row		:.Col = C_W1 : datW1=Cdate(.Text)
						.Row=Row-1	:.Col = C_W1
						If .text<>"" Then 
							datW1_DOWN =Cdate(.text)
							If datW1 < datW1_DOWN Then ' �Ʒ��ຸ�� ��¥�� ������ ���� 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '��: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
						
					End If
					
					.Col = C_W3	: dblW3 = UNICDbl(.text)
					.Col = C_W4	: dblW4 = UNICDbl(.text)
					
					If dblW3 > 0 Or dblW4 > 0 Then
					
						' 2. �ϼ� ���� ��û 
						Call SetW6(Index, Row)					
		
						' 2. ���� ���� 
						Call SetW7(Index, Row)	
					End If
				End If
	
			Case C_W3, C_W4		' ��/�뺯 

				' 1. ���� �����Ϻ��� �Է�üũ 
				.Col = C_W1		: .Row = Row	
				If .Text = "" Then	
					Call DisplayMsgBox("W30002", parent.VB_INFORMATION, "X", "X")           '��: "���ڸ� ���� �Է��Ͻʽÿ�.."
					.Col = Col	: .text =""
					Exit Sub				
				End If
							
				' 2. ���� üũ 
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.text)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.text = 0
				End If

				' 3. �÷� �հ� ��� 
				'dblSum = FncSumSheet(lgvspdData(Index), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 

				' 4. ������ ����Ÿ üũ 
				If Row > 1 Then
					.Row = Row -1	: .Col = C_W5	
					If .text = "" Then
						Call DisplayMsgBox("W30003", parent.VB_INFORMATION, "X", "X")           '��: "������ ��/�뺯�� ���� �Է��Ͻʽÿ�."
						Exit Sub
					End If
				End If
					
				' 5. �ܾ� ��� 
				
				.Row = Row
				.Col = C_W3	: dblW3 = UNICDbl(.text)
				.Col = C_W4	: dblW4 = UNICDbl(.text)
				
				dblW5 =   dblW4-dblW3	' �ܾ� 
				
				' 4.1 ù������ üũ 
				If Row - 1 = 0 Then
					.Col = C_W5 : .text = dblW5
				Else
					' ù���� �ƴҶ� 
					.Row = Row -1
					.Col = C_W5	: dblW5_UP = UNICDbl(.text)	' ���� �ܾ� 
					.Row = Row
					.Col = C_W5	: .text = dblW5_UP + dblW5		' ���� �ܾ�+������ �ܾ� 
				End If
				
				' �ڽ��� ����Ǹ� �ٸ� ȣ���ؾ� �� �κ��� �Ʒ��� ��� 
				Call SetW5(Index, Row)
				
				' �ϼ� ���� ��û 
				Call SetW6(Index, Row)
				
				Call SetW7(Index, Row)	' ���� ��� 
		End Select
	
	Else
		Select Case Col
			Case C_W1	' ������ ����� 
				If lgFISC_END_DT = "" Then
					MsgBox "���� �⺻������ ������� �������� ����ֽ��ϴ�"
					Exit Sub
				Else
					iMaxRows = .MaxRows
					
					' 1-1. ���� �Է��� �������� �������� �����ຸ�� ũ�� ������ ����Ų��.
					If Row + 1 <> iMaxRows Then
						.Row = Row		: .Col = C_W1	: datW1 = CDate(.Text)
						
						' 1.1 �Ʒ����� ���� ��� 
						.Row = Row+1	: .Col = C_W1	
						If .Text <> "" Then
							datW1_DOWN = CDate(.Text)

							If datW1 > datW1_DOWN Then ' �Ʒ��ຸ�� ��¥�� ���ĸ� ���� 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '��: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
					'1-2.�����Է��� �������� �������� �����ຸ�� ������ ����
					ElseIf Row-1 <> 0 Then 
						.Row=Row		:.Col = C_W1 : datW1=Cdate(.Text)
						.Row=Row-1	:.Col = C_W1
						If .text<>"" Then 
							datW1_DOWN =Cdate(.text)
							If datW1 < datW1_DOWN Then ' �Ʒ��ຸ�� ��¥�� ������ ���� 
								Call DisplayMsgBox("WC0016", "X", "X", "X")           '��: "Will you destory previous data"
								.Row=Row : .Col =C_W1:	.TEXT=""
								Exit Sub						
							End If
						End If
						
					End If
					
					.Col = C_W3	: dblW3 = UNICDbl(.text)
					.Col = C_W4	: dblW4 = UNICDbl(.text)
					
					If dblW3 > 0 Or dblW4 > 0 Then
					
						' 2. �ϼ� ���� ��û 
						Call SetW6(Index, Row)					
		
						' 2. ���� ���� 
						Call SetW7(Index, Row)	
					End If
				End If
	
			Case C_W3, C_W4		' ��/�뺯 

				' 1. ���� �����Ϻ��� �Է�üũ 
				.Col = C_W1		: .Row = Row	
				If .Text = "" Then	
					Call DisplayMsgBox("W30002", parent.VB_INFORMATION, "X", "X")           '��: "���ڸ� ���� �Է��Ͻʽÿ�.."
					.Col = Col	: .text = 0
					'Exit Sub				
				End If
							
				' 2. ���� üũ 
				.Col = Col	: .Row = Row	: dblSum = UNICDbl(.text)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "X", "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.text = 0
				End If

				' 3. �÷� �հ� ��� 
				'dblSum = FncSumSheet(lgvspdData(Index), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 

				' 4. ������ ����Ÿ üũ 
				If Row > 1 Then
					.Row = Row -1	: .Col = C_W5	
					If .text = "" Then
						Call DisplayMsgBox("W30003", parent.VB_INFORMATION, "X", "X")           '��: "������ ��/�뺯�� ���� �Է��Ͻʽÿ�."
						Exit Sub
					End If
				End If
					
				' 5. �ܾ� ��� 
				
				.Row = Row
				.Col = C_W3	: dblW3 = UNICDbl(.text)
				.Col = C_W4	: dblW4 = UNICDbl(.text)
					
				dblW5 =  dblW3 - dblW4	' �ܾ� 
				
				' 4.1 ù������ üũ 
				If Row - 1 = 0 Then
					.Col = C_W5 : .text = dblW5
				Else
					' ù���� �ƴҶ� 
					.Row = Row -1
					.Col = C_W5	: dblW5_UP = UNICDbl(.text)	' ���� �ܾ� 
					.Row = Row
					.Col = C_W5	: .Value = dblW5_UP + dblW5		' ���� �ܾ�+������ �ܾ� 
				End If
				
				' �ڽ��� ����Ǹ� �ٸ� ȣ���ؾ� �� �κ��� �Ʒ��� ��� 
				Call SetW5(Index, Row)
				
				' �ϼ� ���� ��û 
				Call SetW6(Index, Row)
				
				Call SetW7(Index, Row)	' ���� ��� 
		End Select
	
	End If
	
	End With
	
End Sub

' �ϼ� ���� 
Sub SetW6(Index, Row)	
	Dim dblW5, dblW6, datW1, datW1_DOWN, dblSum, iRow, blnPrintLast
	
	With lgvspdData(Index)
		blnPrintLast = False
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		For iRow = .MaxRows-1 To 1 Step -1
			.Row = iRow
			.Col = C_W5	: dblW5 = UNICDbl(.text)	' �ܾװ� �б� 
			.Col = C_W1	
		
			If .Text = "" Then		' �ܾ��� ���ų�, �������� �����̸� �հ��� �����Ѵ�.
				
			Else		

				datW1 = CDate(.Text)
		
				If blnPrintLast = False Then	' �������� �ϼ� �����Ѱ�� 
					If frm1.cboREP_TYPE.value = "2" Then
						.Col = C_W6	: .text = DateDiff("d", datW1, DateAdd("m", 6, lgFISC_START_DT)-1)+1
					Else
						.Col = C_W6	: .text = DateDiff("d", datW1, lgFISC_END_DT)+1
					End If
					blnPrintLast = True
				Else
					.Col = C_W1	: .Row = iRow+1	
					
					If .Text <> "" Then	' �����Ҷ�.
						datW1_DOWN = CDate(.Text)	' ���� �������� ���ڸ� ��� 
						.Col = C_W6	: .Row = iRow	: .text = DateDiff("d", datW1,  datW1_DOWN)	
					End If
				End If
			
			ggoSpread.UpdateRow iRow	
			End If
		Next
		
		'dblSum = FncSumSheet(lgvspdData(Index), C_W6, 1, .MaxRows - 1, true, .MaxRows, C_W6, "V")	' �հ� 

		'Call UpdateTotalLIne
		
	End With	

End Sub

Sub UpdateTotalLIne()
	ggoSpread.Source = lgvspdData(lgCurrGrid)
	ggoSpread.UpdateRow lgvspdData(lgCurrGrid).MaxRows
End Sub

' �ܾ� �÷��� ����ɶ� ȣ��� 
Sub SetW5(Index, Row)
	Dim dblSum, dblW3, dblW4, dblW5, dblW6, dblW7, iRow, iMaxRows, dblW5_UP, datW1
	
	With lgvspdData(Index)

		iMaxRows = .MaxRows
		.Row = Row	' ȣ��� ������ Row
		.Col = C_W5	: dblW5 = UNICDbl(.text)	' ���� �������� �ܾ��� ��� 
		.Col = C_W6 : dblW6 = UNICDbl(.text)
		.Col = C_W7 : dblW7 = dblW5 * dblW6	: .text = dblW7	' ���� ���� 
		
		.Row = Row + 1	' ���� �� 
		.Col = C_W3	: dblW3 = UNICDbl(.text)	' ���� 
		.Col = C_W4	: dblW4 = UNICDbl(.text)	' �뺯 
						
		If Row = iMaxRows -1 then 'Or (dblW3 = 0 And dblW4 = 0) Then	' �հ� ���� ���̰ų�, �����࿡ ��/�뺯�� �����̸� 
'			' �Ϸ� �Ͽ����Ƿ� �հ踦 ���Ѵ�.
'			dblSum = FncSumSheet(lgvspdData(Index), C_W5, 1, .MaxRows - 1, true, .MaxRows, C_W5, "V")	' �հ� 
'			Call UpdateTotalLIne
			Exit Sub
		End If

		.Col = C_W5	: dblW5 = dblW5 + (dblW3 - dblW4) 
		.text = dblW5

		'.Col = C_W1 : datW1 = .Text	' ������ 
	End With
	
	' �ڽ��� ����Ǹ� �ٸ� ȣ���ؾ� �� �κ��� �Ʒ��� ��� 
	Call SetW5(Index, Row + 1)	' �������� �����Ͽ����Ƿ�, ��� ������ �����Ѵ�. �� 1-10��(�հ�11��)�� ���� ���, 5���� ��ġ�� 6-10����� ���ľ� �Ѵ�.	

End Sub

Sub SetW7(Index, Row)	' ������ ����ɽ� �ؾߵ� �̺�Ʈ 

	' �ڽ��� ����Ǹ� �ٸ� ȣ���ؾ� �� �κ��� �Ʒ��� ��� 
	Dim dblW5, dblW6, iRow, dblSum
	
	With lgvspdData(Index)
		For iRow = 1 To .MaxRows -1 
			.Row = iRow
			.Col = C_W5 : dblW5 = UNICDbl(.text)
			.Col = C_W6 : dblW6 = UNICDbl(.text)
			If dblW5 <> 0 And dblW6 <> 0 Then 
				.Col = C_W7 : .text = dblW5 * dblW6
			End If
		Next
		
		dblSum = FncSumSheet(lgvspdData(Index), C_W7, 1, .MaxRows - 1, true, .MaxRows, C_W7, "V")	' �հ� 
		Call UpdateTotalLIne
	End With
End Sub


Sub SetW14() ' ������� 
	Dim dblW12, dblW13


	With lgvspdData(TYPE_6)
	
		.Col = C_W12 : 	dblW12 = UNICDbl(.text)
		.Col = C_W13 : dblW13 = UNICDbl(.text)
		.Col = C_W14 : .text = dblW12 * dblW13
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
       
'       If lgSortKey = 1 Then
 '          ggoSpread.SSSort Col               'Sort in ascending
  '         lgSortKey = 2
   '    Else
    '       ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
     '      lgSortKey = 1
      ' End If
       
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
    For i = TYPE_1 To TYPE_6
    
		ggoSpread.Source = lgvspdData(i)
		IF ggoSpread.SSDefaultCheck = False Then								  '��: Check contents area
			Exit Function
		End If
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
		End If
		
	Next
	
	If blnChange = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

    For i = TYPE_1 To TYPE_5
		' ���� '1 �� �ܾ��� ���� < 0 ������ ����. WC0006
		With lgvspdData(i)
		If .MaxRows > 0 Then
			.Row = .MaxRows : .Col = C_W5
			If UNICDbl(.text) < 0 Then
				Select Case i
					Case TYPE_1
						sMsg = "1. ���������δ���� ����"
					Case TYPE_2
						sMsg = "2. �������������� ����"
					Case TYPE_3
						sMsg = "3. Ÿ�����ֽ��� ����"
					Case TYPE_4A
						sMsg = "4. �����ޱ��� ����" & vbCrLf & " ��. �����ޱ��� ����"
					Case TYPE_4B
						sMsg = "4. �����ޱ��� ����" & vbCrLf & " ��. �������� ����"
					Case TYPE_5
						sMsg = "5. ��Ÿ ����"
					Case TYPE_6
						sMsg = "6. �ڱ��ں� �������"
				End Select
				sMsg = sMsg & "�� (5)�ܾ� �հ� "
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, sMsg, "X")                          
				Exit Function
			End If
		End If
		End With
	Next
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True         
    Set gActiveElement = document.ActiveElement                                                 
    
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

			lgvspdData(lgCurrGrid).Col = C_W21
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
	
	If lgCurrGrid = TYPE_6 Then Exit Function
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
	
'		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W3, 1, .MaxRows - 1, true, .MaxRows, C_W3, "V")	' �հ� 
'		dblSum = FncSumSheet(lgvspdData(lgCurrGrid), C_W4, 1, .MaxRows - 1, true, .MaxRows, C_W4, "V")	' �հ� 
	
		' �ڽ��� ����Ǹ� �ٸ� ȣ���ؾ� �� �κ��� �Ʒ��� ��� 
		Call SetW5(lgCurrGrid, 1)
					
		' �ϼ� ���� ��û 
		Call SetW6(lgCurrGrid, 1)
					
		Call SetW7(lgCurrGrid, 1)	' ���� ��� 
	End With
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
   	 
	If lgCurrGrid = TYPE_6 Then	Exit Function	' 6�� �׸���� �߰��Ҽ� ����.
	
	
	With lgvspdData(lgCurrGrid)	' ��Ŀ���� �׸��� 
		
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		
		iRow = .ActiveRow
		lgvspdData(lgCurrGrid).ReDraw = False
				
		If .MaxRows = 0 Then	' ù InsertRow�� 1��+�հ��� 

			iRow = 1
			ggoSpread.InsertRow ,2
			Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
			.Row = iRow		
			.Col = C_SEQ_NO : .Text = iRow	
			.Col = C_W_TYPE : .Text = lgCurrGrid
			
			If lgCurrGrid = TYPE_6 Then
				ggoSpread.SpreadLock C_W10, iRow, C_W14, iRow
				
			Else	
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
				.Col = C_W_TYPE : .Text = lgCurrGrid	
				.Col = C_W1		: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
						
				ggoSpread.SpreadLock C_W1, iRow, C_W7, iRow
			End If
						
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
	
	Call CheckW7Status(lgCurrGrid)	' ������ ���� üũ 

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function

' -- TYPE_4A, TYPE_4B �׸����� ���� Enable/Disable üũ 
Function CheckW7Status(Index)

	If Index = TYPE_4A Or Index = TYPE_4B Then
	
		With lgvspdData(Index)
	
		ggoSpread.Source = lgvspdData(Index)

		If lgvspdData(Index).MaxRows > 1 Then
			ggoSpread.SpreadLock C_W7, .MaxRows, C_W7, .MaxRows
		Else
			ggoSpread.SpreadUnLock C_W7, .MaxRows, C_W7, .MaxRows
		End If
	
		End With
	End If
End Function

' GetREF ���� ���� �����µ� ȣ��� 
Function InsertTotalLine(Index)
	With lgvspdData(Index)
	
	ggoSpread.Source = lgvspdData(Index)
	
	If .MaxRows = 0 Then	' ���� �߰� 
		ggoSpread.InsertRow ,1
		
		.Row = 1
		.Col = C_SEQ_NO : .Text = SUM_SEQ_NO
		.Col = C_W_TYPE : .Text = Index	
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
		.Col = C_W_TYPE : .Value = lgCurrGrid	' ���� �׸��� ��ȣ 
		MaxSpreadVal lgvspdData(lgCurrGrid), C_SEQ_NO, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(lgCurrGrid), C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO : .text = iSeqNo : iSeqNo = iSeqNo + 1
			.Col = C_W_TYPE : .text = lgCurrGrid	' ���� �׸��� ��ȣ 
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

	If lgCurrGrid = TYPE_6 Then Exit Function
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
        strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	If lgvspdData(TYPE_1).MaxRows > 0 Or _
		lgvspdData(TYPE_2).MaxRows > 0 Or _
		lgvspdData(TYPE_3).MaxRows > 0 Or _
		lgvspdData(TYPE_4A).MaxRows > 0 Or _
		lgvspdData(TYPE_4B).MaxRows > 0 Or _
		lgvspdData(TYPE_5).MaxRows > 0 Or _
		lgvspdData(TYPE_6).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		    
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
			'Call SetSpreadLock(TYPE_1)
			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1111111100011111")										<%'��ư ���� ���� %>
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W7, -1
			Call SetToolbar("1110000000011111")										<%'��ư ���� ���� %>
		End If
	
	Else
		Call SetToolbar("1100110100111111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
	Call ClickTab1()
	lgvspdData(lgCurrGrid).focus			
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
    
    For i = TYPE_1 To TYPE_6	' ��ü �׸��� ���� 
    
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
					For lCol = C_SEQ_NO To lMaxCols
						If lCol=C_W1  Then 
							.Col = lCol : 
							If  trim(.Text)="��" or trim(.text)="Blank" Then 
								strVal = strVal & Trim("2999-12-31") &  Parent.gColSep
							Else
								strVal = strVal & Trim(.Text) &  Parent.gColSep
							End IF
						Else
							
						.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
						End If
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
			Next
		
		End With

	Next


	Frm1.txtSpread.value      = strVal '& strVal
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	
	For iRow = TYPE_1 To TYPE_6
		lgvspdData(lgCurrGrid).MaxRows = 0
		ggoSpread.Source = lgvspdData(lgCurrGrid)
		ggoSpread.ClearSpreadData
	Next
	
	lgvspdData(lgCurrGrid).MaxRows = 1
	
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()" width=200>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������� �ε���/������ ����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Ÿ���� �ֽ�/�����ޱ��� ����</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��Ÿ/�ڱ��ں� �������</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w3127ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=15%>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;1. �������� �ε����� ����							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData0_vspdData0.js'></script>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;2. �������� ������ ����							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData1_vspdData1.js'></script>
										</TD>
									</TR>
								</TABLE>
								</DIV>
						
								<DIV ID="TabDiv" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;3. Ÿ�����ֽ��� ����							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="30%">
											<script language =javascript src='./js/w3127ma1_vspdData2_vspdData2.js'></script>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;4. �����ޱ��� ����							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="60%">
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD HEIGHT="10">&nbsp;��. �����ޱݵ��� ����							
												</TD>
											</TR>
											<TR>
												<TD HEIGHT="45%">
													<script language =javascript src='./js/w3127ma1_vspdData3_vspdData3.js'></script>
												</TD>
											</TR>
											<TR>
												<TD HEIGHT="10">&nbsp;��. �����ݵ��� ����							
												</TD>
											</TR>
											<TR>
												<TD HEIGHT="45%">
													<script language =javascript src='./js/w3127ma1_vspdData4_vspdData4.js'></script>
												</TD>
											</TR>								
											</TABLE>
										</TD>
									</TR>
								</TABLE>
								</DIV>

								<DIV ID="TabDiv" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;5. ��Ÿ����							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData5_vspdData5.js'></script>
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="10">&nbsp;6. �ڱ��ں� �������							
										</TD>
									</TR>
									<TR>
										<TD HEIGHT="45%">
											<script language =javascript src='./js/w3127ma1_vspdData6_vspdData6.js'></script>
										</TD>
									</TR>
								</TABLE>
								</DIV>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>���������Ǻε����� ����</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>���������ǵ����� ����</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check3" ><LABEL FOR="prt_check3"><����>Ÿ�����ֽ��� ����</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check4" ><LABEL FOR="prt_check4"><����>�����ޱݵ��� ����</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check5" ><LABEL FOR="prt_check5"><����>�����ݵ��� ����</LABEL>&nbsp;
				        <INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check6" ><LABEL FOR="prt_check6"><����>��Ÿ ����</LABEL>&nbsp;
				        
				        </TD>
				 
                </TR>
			
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

