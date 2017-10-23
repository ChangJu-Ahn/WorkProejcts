
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���Աݾ����� 
'*  3. Program ID           : W1111MA1
'*  4. Program Name         : W1111MA1.asp
'*  5. Program Desc         : ��16-2ȣ ���Թ��� ���� 
'*  6. Modified date(First) : 2004/12/30
'*  7. Modified date(Last)  : 2004/12/30
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

Const BIZ_MNU_ID = "W2101MA1"
Const BIZ_PGM_ID = "w2101mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID = "w2101mb2.asp"
Const EBR_RPT_ID		= "W2101OA1"
Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.


Dim C_SEQ_NO1
Dim C_W7
Dim C_W8
Dim C_W8_NM
Dim C_W9
Dim C_W10
Dim C_W11
Dim C_W12
Dim C_W13

Dim C_SEQ_NO2
Dim C_W14
Dim C_W15
Dim C_W16
Dim C_W17
Dim C_W18
Dim C_W19
Dim C_W20
Dim C_W21
Dim C_HEAD_SEQ_NO1

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgCurrGrid, lgvspdData(1)
Dim	lgTB_26_AMT, lgTB_3_AMT	

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()

	C_SEQ_NO1	= 1	' -- 1�� �׸��� 
    C_W7		= 2
    C_W8		= 3
    C_W8_NM		= 4
    C_W9		= 5
    C_W10		= 6
    C_W11		= 7
    C_W12		= 8
    C_W13		= 9	
 
 	C_SEQ_NO2	= 1  ' -- 2�� �׸��� 
    C_W14		= 2 
    C_W15		= 3
    C_W16		= 4
    C_W17		= 5
    C_W18		= 6
    C_W19		= 7
    C_W20		= 8
    C_W21		= 9
    C_HEAD_SEQ_NO1 = 10
    
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
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
    
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1003' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboW2 ,lgF0  ,lgF1  ,Chr(11))    
End Sub

Sub InitSpreadSheet()
	Dim ret
	
    Call initSpreadPosVariables()  

	Set lgvspdData(TYPE_1) = frm1.vspdData0
	Set lgvspdData(TYPE_2) = frm1.vspdData1

	' 1�� �׸��� 
	With lgvspdData(TYPE_1)
	
	ggoSpread.Source = lgvspdData(TYPE_1)	
   'patch version
    ggoSpread.Spreadinit "V20041222" & TYPE_1,,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_W13 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols									'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    ' 
    Call AppendNumberPlace("6","3","1")

    ggoSpread.SSSetEdit		C_SEQ_NO1,	"����"		, 10,,,15,1	' �����÷� 
	ggoSpread.SSSetEdit		C_W7,		"(7)���θ�"	, 20,,,50,1	
    ggoSpread.SSSetCombo	C_W8,		"(8)����"		, 10
    ggoSpread.SSSetCombo	C_W8_NM,	"(8)����"		, 10	
    ggoSpread.SSSetEdit		C_W9,		"(9)����ڵ�Ϲ�ȣ", 20,,,20,1
    ggoSpread.SSSetEdit		C_W10,		"(10)������"	, 30,,,250,1
    ggoSpread.SSSetEdit		C_W11,		"(11)��ǥ"		, 15,,,50,1
    ggoSpread.SSSetFloat	C_W12,		"(12)�����ֽ��Ѽ�" , 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0",""  
    ggoSpread.SSSetFloat	C_W13,		"(13)������"	, 10, "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","100" 

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_W8,C_W8,True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO1,C_SEQ_NO1,True)
				
	Call InitSpreadComboBox()
	
	.ReDraw = true
	
    'Call SetSpreadLock 
    
    End With

 	' -----  2�� �׸��� 
	With lgvspdData(TYPE_2)
	
	ggoSpread.Source = lgvspdData(TYPE_2)	
   'patch version
    ggoSpread.Spreadinit "V20041222_2" & TYPE_2,,parent.gAllowDragDropSpread    
    
	.ReDraw = false
    
    .MaxCols = C_HEAD_SEQ_NO1 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols									'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

	'����� 2�ٷ�    
    .ColHeaderRows = 2
    'Call AppendNumberPlace("6","3","2")

    ggoSpread.SSSetEdit		C_SEQ_NO2,	"����", 10,,,15,1	' �����÷� 
	ggoSpread.SSSetEdit		C_W14,		"(14)��ȸ�� �Ǵ�" & vbCrLf & "�������� ���θ�", 20,,,50,1
    ggoSpread.SSSetFloat	C_W15,		"(15)���ݾ�",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
    ggoSpread.SSSetCombo	C_W16,		"(16)�ͱݺһ�����" , 10, 1
	ggoSpread.SSSetFloat	C_W17,		"(17)�ͱݺһ��� ���ݾ�",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0","" 
    ggoSpread.SSSetFloat	C_W18,		"(18)�Ұ�",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0",""
    ggoSpread.SSSetFloat	C_W19,		"(19)��$18��2 (��$18��3)",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0","" 
    ggoSpread.SSSetFloat	C_W20,		"(20)��$18��2 ��1����4ȣ" ,15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0",""
    ggoSpread.SSSetFloat	C_W21,		"(21)�ͱ� �һ��Ծ�" & vbCrLf & "(17-18)" , 15,Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"","0",""
	ggoSpread.SSSetEdit		C_HEAD_SEQ_NO1,	"�������", 10,,,15,1	' �����÷� 
 
    ret = .AddCellSpan(1, -1000, 1, 2)
    ret = .AddCellSpan(2, -1000, 1, 2)
    ret = .AddCellSpan(3, -1000, 1, 2)
    ret = .AddCellSpan(4, -1000, 1, 2)
    ret = .AddCellSpan(5, -1000, 1, 2)
    ret = .AddCellSpan(6, -1000, 3, 1)
    ret = .AddCellSpan(9, -1000, 1, 2) 
    ret = .AddCellSpan(10, -1000, 1, 2) 
    
    ' ù��° ��� ��� ���� 
	.Row = -1000
	.Col = 6
	.Text = "�ͱݺһ��������ݾ�"

	' �ι�° ��� ��� ���� 
	.Row = -999
	.Col = 6
	.Text = "(18)�Ұ�"
	.Col = 7
	.Text = "(19)��$18��2" & vbCrLf & "(��$18��3)"
	.Col = 8
	.Text = "(20)��$18��2" & vbCrLf & "��1����4ȣ"
	.rowheight(-999) = 20	' ���� ������ 
	
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO2,C_SEQ_NO2,True)
	Call ggoSpread.SSSetColHidden(C_HEAD_SEQ_NO1,C_HEAD_SEQ_NO1,True)
				
	Call InitSpreadComboBox2()
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
       
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()

    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx

	' ȸ�籸�� 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " (MAJOR_CD='W1004') ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = lgvspdData(TYPE_1)
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_W8
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_W8_NM
	End If
		  
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

End Sub

Sub InitSpreadComboBox2()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx, i, sVal, sVal2, sVal3
    
    ggoSpread.Source = lgvspdData(TYPE_2)
    
    sVal = "100" & vbTab & "90" & vbTab & "60" & vbTab & "50" & vbTab & "30"

	ggoSpread.SetCombo sVal, C_W16

End Sub

Sub SetSpreadLock()

    lgvspdData(TYPE_2).ReDraw = False
    ggoSpread.Source = lgvspdData(TYPE_2)	

    ggoSpread.SpreadLock C_W17, -1, C_W17
    ggoSpread.SpreadLock C_W18, -1, C_W18
    ggoSpread.SpreadLock C_W21, -1, C_W21
	ggoSpread.SSSetRequired C_W14, -1, -1
	ggoSpread.SSSetRequired C_W15, -1, -1
	ggoSpread.SSSetRequired C_W16, -1, -1
	'ggoSpread.SSSetRequired C_W19, -1, -1
	ggoSpread.SpreadLock C_W14, lgvspdData(TYPE_2).MaxRows, C_W21
	lgvspdData(TYPE_2).ReDraw = True
	
	lgvspdData(TYPE_1).ReDraw = False
    ggoSpread.Source = lgvspdData(TYPE_1)
        
	ggoSpread.SSSetRequired C_W7, -1, -1
	ggoSpread.SSSetRequired C_W8, -1, -1
	ggoSpread.SSSetRequired C_W8_NM, -1, -1
	ggoSpread.SSSetRequired C_W9, -1, -1
	ggoSpread.SSSetRequired C_W10, -1, -1
	ggoSpread.SSSetRequired C_W11, -1, -1
	ggoSpread.SSSetRequired C_W12, -1, -1
	ggoSpread.SSSetRequired C_W13, -1, -1
	lgvspdData(TYPE_1).ReDraw = True
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)
    With lgvspdData(pType)

	If pType = TYPE_1 Then
		.ReDraw = False
 
		ggoSpread.Source = lgvspdData(pType)
	
  		ggoSpread.SSSetRequired C_W7, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W8, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W8_NM, pvStartRow, pvEndRow 		
 		ggoSpread.SSSetRequired C_W9, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W10, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W11, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W12, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W13, pvStartRow, pvEndRow
		    
		.ReDraw = True
    Else
		.ReDraw = False
 
		ggoSpread.Source = lgvspdData(pType)
	
		ggoSpread.SSSetRequired C_W14, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W15, pvStartRow, pvEndRow
 		ggoSpread.SSSetRequired C_W16, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_W17, pvStartRow, pvEndRow
  		ggoSpread.SSSetProtected C_W18, pvStartRow, pvEndRow
  		'ggoSpread.SSSetRequired C_W19, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected C_W21, pvStartRow, pvEndRow
  		
		.ReDraw = True    
    End If
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case pvSpdNo
       Case TYPE_1
            ggoSpread.Source = lgvspdData(TYPE_1)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_SEQ_NO1	= iCurColumnPos(1)	' -- 1�� �׸��� 
			C_W7		= iCurColumnPos(2)
			C_W8		= iCurColumnPos(3)
			C_W8_NM		= iCurColumnPos(4)
			C_W9		= iCurColumnPos(5)
			C_W10		= iCurColumnPos(6)
			C_W11		= iCurColumnPos(7)
			C_W12		= iCurColumnPos(8)
			C_W13		= iCurColumnPos(9)	
			
 		Case TYPE_2
 			ggoSpread.Source = lgvspdData(TYPE_2)
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
  			C_SEQ_NO2	= iCurColumnPos(1)  ' -- 2�� �׸��� 
			C_W14		= iCurColumnPos(2) 
			C_W15		= iCurColumnPos(3)
			C_W16		= iCurColumnPos(4)
			C_W17		= iCurColumnPos(5)
			C_W18		= iCurColumnPos(6)
			C_W19		= iCurColumnPos(7)
			C_W20		= iCurColumnPos(8)
			C_W21		= iCurColumnPos(9)
			C_HEAD_SEQ_NO1 = iCurColumnPos(10)
               
    End Select    
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

    'Call ggoOper.ClearField(Document, "2")	
    ggoSpread.Source = lgvspdData(TYPE_1)
    ggoSpread.ClearSpreadData
    
    ggoSpread.Source = lgvspdData(TYPE_2)
    ggoSpread.ClearSpreadData
    
    Call InitVariables 
    			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
	' 2. ��������ǥ�� �ڻ��Ѱ�, ��ä�Ѱ�-�����޹��μ�, �ں���+�����޹��μ�+�ֽĹ����ʰ���+��������-�ֽĹ�����������-�������� �������� 
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
			
			'Call vspdData_Change(TYPE_2, C_W15, iRow)
			'Call vspdData_Change(TYPE_2, C_W19, iRow)
			'.Col = C_W15	: dblVal(C_W15) = UNICDbl(.value)
			'.Col = C_W16	: dblVal(C_W16) = UNICDbl(.Text)
			'If dblVal(C_W15) = 0 And dblVal(C_W16) = 0 Then Exit Sub
			'dblVal(C_W17) = dblVal(C_W15) * (dblVal(C_W16) * 0.01)
			'.Col = C_W17	: .value = dblVal(C_W17)		
		
		Next
	
	
	End With
End Sub

'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110100011111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

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

Sub InitData()
	Dim sCoCd , sFiscYear, sRepType
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	lgCurrGrid = TYPE_1
	
	Dim iRet
	sCoCd		= "<%=wgCO_CD%>"
	sFiscYear	= "<%=wgFISC_YEAR%>"
	sRepType	= "<%=wgREP_TYPE%>"
	' �������� ��� 
	iRet = CommonQueryRs("W1, W2, W3, W4, W5, W6"," dbo.ufn_TB_16_2_GetRef4('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If iRet = True Then
		With frm1
		.txtW1.value = Replace(lgF0, Chr(11), "")
		.cboW2.value = Replace(lgF1, Chr(11), "")
		.txtW3.value = Replace(lgF2, Chr(11), "")
		.txtW4.value = Replace(lgF3, Chr(11), "")
		.txtW5.value = Replace(lgF4, Chr(11), "")
		.txtW6.value = Replace(lgF5, Chr(11), "")
		
		Call cboW2_onChange
		End With
    End If
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
Sub vspdData_ComboSelChange(Byval pType, ByVal Col, ByVal Row)
	Dim iIdx, dblVal(30)
	
	lgBlnFlgChgValue = True
	With lgvspdData(pType)
	
		Select Case pType
			Case TYPE_2
				If Col = C_W16 Then
					.Row = Row
					.Col = C_W15	: dblVal(C_W15) = UNICDbl(.value)
					.Col = C_W16	: dblVal(C_W16) = UNICDbl(.Text)
					If dblVal(C_W15) = 0 And dblVal(C_W16) = 0 Then Exit Sub
					dblVal(C_W17) = dblVal(C_W15) * (dblVal(C_W16) * 0.01)
					.Col = C_W17	: .value = dblVal(C_W17)
				End If
			
			Case TYPE_1
				If Col = C_W8_NM Then
					.Row = Row
					.Col = C_W8_NM	: iIdx = UNICDbl(.Value)
					.Col = C_W8		: .Value = iIdx
				End If
		End Select
	End With
End Sub
	
Sub vspdData_Change(Byval pType, ByVal Col , ByVal Row )
    lgvspdData(pType).Row = Row
    lgvspdData(pType).Col = Col

    If lgvspdData(TYPE_1).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(TYPE_1).text) < UNICDbl(lgvspdData(TYPE_1).TypeFloatMin) Then
         lgvspdData(TYPE_1).text = lgvspdData(TYPE_1).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(pType)
    ggoSpread.UpdateRow Row

	If pType = TYPE_1 Then Exit Sub
	
	Dim dblVal(30)
	
	With lgvspdData(TYPE_2)
	
		Select Case Col
			Case C_W15, C_W16	' W17��� 
				.Row = Row
				.Col = C_W15	: dblVal(C_W15) = UNICDbl(.value)
				.Col = C_W16	: dblVal(C_W16) = UNICDbl(.Text)
				If dblVal(C_W15) = 0 And dblVal(C_W16) = 0 Then Exit Sub
				dblVal(C_W17) = dblVal(C_W15) * (dblVal(C_W16) * 0.01)
				.Col = C_W17	: .value = dblVal(C_W17)
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W15, 1, .MaxRows - 1, true, .MaxRows, C_W15, "V")	' �հ� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W17, 1, .MaxRows - 1, true, .MaxRows, C_W17, "V")	' �հ� 
			
				' C_W17 �� ���������Ƿ� �ش��̺�Ʈ�� �߻��� �ش�.
				Call vspdData_Change(pType, C_W17, Row)
				
			Case C_W17, C_W18
				.Row = Row
				.Col = C_W17	: dblVal(C_W17) = UNICDbl(.value)
				.Col = C_W18	: dblVal(C_W18) = UNICDbl(.value)
				
				dblVal(C_W21) = dblVal(C_W17) - dblVal(C_W18)
				.Col = C_W21	: .value = dblVal(C_W21)
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W21, 1, .MaxRows - 1, true, .MaxRows, C_W21, "V")	' �հ� 
			
				' C_W21 �� ���������Ƿ� �ش��̺�Ʈ�� �߻��� �ش�.
				Call vspdData_Change(pType, C_W21, Row)
				
			Case C_W19, C_W20
				.Row = Row
				.Col = C_W19	: dblVal(C_W19) = UNICDbl(.value)
				.Col = C_W20	: dblVal(C_W20) = UNICDbl(.value)
				
				dblVal(C_W18) = dblVal(C_W19) + dblVal(C_W20)
				.Col = C_W18	: .value = dblVal(C_W18)
				
				Call FncSumSheet(lgvspdData(lgCurrGrid), Col, 1, .MaxRows - 1, true, .MaxRows, Col, "V")	' �հ� 
				Call FncSumSheet(lgvspdData(lgCurrGrid), C_W18, 1, .MaxRows - 1, true, .MaxRows, C_W18, "V")	' �հ�		
			
				' C_W18 �� ���������Ƿ� �ش��̺�Ʈ�� �߻��� �ش�.
				Call vspdData_Change(pType, C_W18, Row)
				
			Case C_W14 
				' -- �׸���1�� (7)���θ��� �˻��� SEQ_NO�� C_HEAD_SEQ_NO1�� �ִ´� 
				.Row = Row
				.Col = Col
				dblVal(C_HEAD_SEQ_NO1) = SearchW7(.Text)
				
				If dblVal(C_HEAD_SEQ_NO1) = -1 Then
					Call DisplayMsgBox("W20002", parent.VB_INFORMATION, .Text, "X")
					.Value = ""
					Exit Sub
				End If
				.Col = C_HEAD_SEQ_NO1	: .value = dblVal(C_HEAD_SEQ_NO1)
		End Select
		
		ggoSpread.UpdateRow .MaxRows
	End With
End Sub

Sub ReClacGrid2Sum()
	Dim iCol
	With lgvspdData(TYPE_2)
	For iCol = C_W15 To C_W21
		If iCol <> C_W16 Then
			Call FncSumSheet(lgvspdData(lgCurrGrid), iCol, 1, .MaxRows - 1, true, .MaxRows, iCol, "V")	' �հ� 
		End If
	Next
	End With
End Sub

Function SearchW7(Byval pCoNm)
	Dim iMaxRows, iRow
	
	With lgvspdData(TYPE_1)
		iMaxRows = .MaxRows
		
		For iRow = 1 To iMaxRows
			.Row = iRow
			.Col = C_W7
			If UCase(.Text) = UCase(pCoNm) Then
				.Col = C_SEQ_NO1
				SearchW7 = UNICDbl(.Value)
				Exit Function
			End If
		Next

	End With
	SearchW7 = -1	' --������� 
End Function

Sub vspdData_Click(Byval pType, ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

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

	lgvspdData(TYPE_1).Row = Row
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
    
    if lgvspdData(TYPE_1).MaxRows < NewTop + VisibleRowCnt(lgvspdData(TYPE_1),NewTop) Then	           
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
    ggoSpread.Source = lgvspdData(TYPE_1)
    If ggoSpread.SSCheckChange = True Then
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
    
    
    Dim blnChange
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    

<%  '-----------------------
    'Precheck area
    '----------------------- %> 

    If Not chkField(Document, "2") Then								<%'This function check indispensable field%>
       Exit Function
    End If    
    
    ggoSpread.Source = lgvspdData(TYPE_1)
    If ggoSpread.SSCheckChange = False Then
        blnChange = False
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If    
	
    ggoSpread.Source = lgvspdData(TYPE_2)
    If ggoSpread.SSCheckChange = False Then
        blnChange = False
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If  

    If lgBlnFlgChgValue = False  Then
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

    Call SetToolbar("1100110100011111")

	Call InitData
	
	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If lgvspdData(lgCurrGrid).MaxRows = 0 Then
       Exit Function
    End If
 
	With lgvspdData(lgCurrGrid)

		.focus
		.ReDraw = False
				
		Select Case lgCurrGrid
			Case TYPE_1
				ggoSpread.Source = lgvspdData(TYPE_1)
		
				ggoSpread.CopyRow
				SetSpreadColor .ActiveRow, .ActiveRow
				MaxSpreadVal lgvspdData(TYPE_1), C_SEQ_NO1, iRow

			Case TYPE_2
				ggoSpread.Source = lgvspdData(TYPE_2)
		
				ggoSpread.CopyRow
				SetSpreadColor .ActiveRow, .ActiveRow
				MaxSpreadVal lgvspdData(TYPE_2), C_SEQ_NO2, iRow
		End Select
	
		.ReDraw = True
	
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    ggoSpread.Source = lgvspdData(TYPE_1)	
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
		lgvspdData(lgCurrGrid).ReDraw = False
					
		If .MaxRows = 0 Then	' ù InsertRow�� 1��+�հ��� 

			iRow = 1			
			If lgCurrGrid = TYPE_1 Then
				ggoSpread.InsertRow , 1
				Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
				
				.Col = C_SEQ_NO1 : .Text = iRow	
			Else
				ggoSpread.InsertRow , 2
				Call SetSpreadColor(lgCurrGrid, iRow, iRow+1) 
				
				.Col = C_SEQ_NO2 : .Text = iRow	
			
				iRow = 2		: .Row = iRow
				.Col = C_SEQ_NO2: .Text = SUM_SEQ_NO	
				.Col = C_W14	: .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				.Col = C_W16	: .CellType = 1	: .TypeHAlign = 1
				ggoSpread.SpreadLock C_W14, iRow, .MaxCols-1, iRow
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
		MaxSpreadVal lgvspdData(Index), C_SEQ_NO1, iRow
	Else
		iSeqNo = MaxSpreadVal(lgvspdData(Index), C_SEQ_NO2, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
		
		For i = iRow to iRow + iAddRows -1
			.Row = i
			.Col = C_SEQ_NO2 : .Value = iSeqNo : iSeqNo = iSeqNo + 1
		Next
	End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows
	
	If lgCurrGrid = 1	Then
		With lgvspdData(TYPE_1) 
			.focus
			ggoSpread.Source = lgvspdData(TYPE_1) 
			lDelRows = ggoSpread.DeleteRow
		End With
    Else
		With lgvspdData(TYPE_2) 
			.focus
			ggoSpread.Source = lgvspdData(TYPE_2)
			lDelRows = ggoSpread.DeleteRow
		End With    
    End If
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
	
    ggoSpread.Source = lgvspdData(TYPE_1)	
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
	
	If lgvspdData(TYPE_1).MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		
		Call SetSpreadLock()
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1101111100011111")										<%'��ư ���� ���� %>
			
		Else
		
			'ggoSpread.SpreadLock	C_W1, -1, C_W5, -1
			Call SetToolbar("1100000000011111")										<%'��ư ���� ���� %>
		End If
	Else
		Call SetToolbar("1100110100011111")										<%'��ư ���� ���� %>
	End If
	
	lgvspdData(TYPE_1).focus			
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
    Dim lRow, lCol, lMaxRows, lMaxCols , i 
 
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
			For lRow = 1 To lMaxRows
    
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
					For lCol = 1 To lMaxCols
						.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
					Next
					strVal = strVal & Trim(.Text) &  Parent.gRowSep
			  End If  
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
						<a href="vbscript:GetRef">��系�� �ҷ�����</A> | <A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w2101ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
                                <TD WIDTH="100%"> 1. ����ȸ�� �Ǵ� ���ڹ��� ��Ȳ</TD>
                            </TR>
                            <TR HEIGHT=10>
								<TD WIDTH="100%">
                                 <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
								     <TR>
								         <TD CLASS="TD51" width="17%" ALIGN=CENTER>(1)���θ�</TD>
								         <TD CLASS="TD51" width="10%" ALIGN=CENTER>(2)����</TD>
								         <TD CLASS="TD51" width="17%" ALIGN=CENTER>(3)����ڵ�Ϲ�ȣ</TD>
								         <TD CLASS="TD51" width="25%" ALIGN=CENTER>(4)������</TD>
								         <TD CLASS="TD51" width="15%" ALIGN=CENTER>(5)��ǥ�ڼ���</TD>
								         <TD CLASS="TD51" width="10%" ALIGN=CENTER>(6)���� ����</TD>
								    </TR>
								    <TR>
								  		<TD><INPUT NAME="txtW1" MAXLENGTH=25 TYPE="Text" ALT="���θ�" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><SELECT NAME="cboW2" ALT="����" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></SELECT><INPUT TYPE=HIDDEN NAME=txtW2_NM></TD>
								  		<TD><INPUT NAME="txtW3" MAXLENGTH=20  TYPE="Text" ALT="����ڵ�Ϲ�ȣ" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><INPUT NAME="txtW4" MAXLENGTH=125  TYPE="Text" ALT="������" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><INPUT NAME="txtW5" MAXLENGTH=25  TYPE="Text" ALT="��ǥ�ڼ���" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								  		<TD><INPUT NAME="txtW6" MAXLENGTH=50  TYPE="Text" ALT="���� ����" style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%"  tag="23"></TD>
								    </TR>
								</table>
								</TD>
							</TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">2. ��ȸ�� �Ǵ� ���� ���޹��� ��Ȳ</TD>
                            </TR>
                            <TR HEIGHT=40%>
                                <TD WIDTH="100%">
								<script language =javascript src='./js/w2101ma1_vaSpread1_vspdData0.js'></script></TD>
                            </TR>
                            <TR HEIGHT=10>
                                <TD WIDTH="100%">3. ���Թ��� �� �ͱݺһ��� �ݾ� ��</TD>
                            </TR>   
                            <TR HEIGHT=60%>
                                <TD WIDTH="100%">
								<script language =javascript src='./js/w2101ma1_vaSpread1_vspdData1.js'></script></TD>
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
				        <TD><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check1" ><LABEL FOR="prt_check1"><����>��ȸ�� �Ǵ� �������޹��� ��Ȳ</LABEL>&nbsp;
							<INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check" TAG="1X" VALUE="YES" ID="prt_check2" ><LABEL FOR="prt_check2"><����>���Թ��� �� �ͱݺһ��� �ݾ� ��</LABEL>&nbsp;
				        
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
<TEXTAREA CLASS="hidden" NAME="txtSpread0" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

