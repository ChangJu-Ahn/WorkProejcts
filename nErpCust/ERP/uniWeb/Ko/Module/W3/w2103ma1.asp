
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ���Աݾ����� 
'*  3. Program ID           : W1113MA1
'*  4. Program Name         : W1113MA1.asp
'*  5. Program Desc         : ���Թ��� �Է� 
'*  6. Modified date(First) : 2004/12/28
'*  7. Modified date(Last)  : 2004/12/28
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "w2103ma1"
Const BIZ_PGM_ID = "w2103mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID = "w2103mb2.asp"

Const TAB1 = 1																	'��: Tab�� ��ġ 
Const TAB2 = 2

Const TYPE_1	= 0		' �׸��� �迭��ȣ �� ����� W_TYPE �÷��� ��. 
Const TYPE_2	= 1		' �� ��Ƽ �׸��� PG������ ���� ���̺��� �ڵ�� �����ȴ�.

Dim C_SEQ_NO
Dim C_DOC_DATE
Dim C_DOC_AMT
Dim C_DEBIT_CREDIT
Dim C_DEBIT_CREDIT_NM
Dim C_SUMMARY_DESC
Dim C_COMPANY_NM
Dim C_STOCK_RATE
Dim C_ACQUIRE_AMT
Dim C_COMPANY_TYPE
Dim C_COMPANY_TYPE_NM
Dim C_HOLDING_TERM
Dim C_JUKSU
Dim C_OWN_RGST_NO
Dim C_CO_ADDR
Dim C_REPRE_NM
Dim C_STOCK_CNT

Dim C_MINOR_NM
Dim C_MINOR_CD
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W3_1
Dim C_W3_2
Dim C_W3_3
Dim C_W3_4
Dim C_W3_5
Dim C_W4

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO			= 1
    C_DOC_DATE			= 2
    C_DOC_AMT			= 3
    C_DEBIT_CREDIT		= 4
    C_DEBIT_CREDIT_NM	= 5
    C_SUMMARY_DESC		= 6
    C_COMPANY_NM		= 7
    C_STOCK_RATE		= 8
    C_ACQUIRE_AMT		= 9
    C_COMPANY_TYPE		= 10
    C_COMPANY_TYPE_NM	= 11
    C_HOLDING_TERM		= 12
    C_JUKSU				= 13
    C_OWN_RGST_NO		= 14
    C_CO_ADDR			= 15
    C_REPRE_NM			= 16
    C_STOCK_CNT			= 17
    
    C_MINOR_NM			= 2
    C_MINOR_CD			= 3
    C_W1				= 4
    C_W2				= 5
    C_W3				= 6
    C_W3_1				= 6
    C_W3_2				= 7
    C_W3_3				= 8
    C_W3_4				= 9
    C_W3_5				= 10
    C_W4				= 11
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
	
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20041222",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_STOCK_CNT + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    Call AppendNumberPlace("6","3","1")

    ggoSpread.SSSetEdit		C_SEQ_NO,		"����", 10,,,100,1
	ggoSpread.SSSetDate		C_DOC_DATE,		"(1)����",			10,		2,		Parent.gDateFormat,	-1
	ggoSpread.SSSetFloat	C_DOC_AMT,		"(2)�ݾ�",		15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec 
    ggoSpread.SSSetCombo	C_DEBIT_CREDIT, "��/��"    , 10
    ggoSpread.SSSetCombo	C_DEBIT_CREDIT_NM, "��/��"    , 10
    ggoSpread.SSSetEdit		C_SUMMARY_DESC, "(3)����", 15,,,200,1
    ggoSpread.SSSetEdit		C_COMPANY_NM,	"(4)ȸ���", 15,,,30
    ggoSpread.SSSetFloat	C_STOCK_RATE,	"(5)������" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    ggoSpread.SSSetFloat	C_ACQUIRE_AMT,	"(6)��氡��" , 15,Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
    ggoSpread.SSSetCombo	C_COMPANY_TYPE, "ȸ�籸��", 10
    ggoSpread.SSSetCombo	C_COMPANY_TYPE_NM, "(7)ȸ�籸��", 10
    ggoSpread.SSSetFloat	C_HOLDING_TERM, "(8)��⺸���Ⱓ", 12, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
    ggoSpread.SSSetFloat	C_JUKSU,	"����" , 15,Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"
    ggoSpread.SSSetEdit		C_OWN_RGST_NO,	"(9)����ڵ�Ϲ�ȣ", 14, 2,,100,1
    ggoSpread.SSSetEdit		C_CO_ADDR,		"(10)������", 20,,,100,1
    ggoSpread.SSSetEdit		C_REPRE_NM,		"(11)��ǥ��", 10,,,100,1
    ggoSpread.SSSetFloat	C_STOCK_CNT,	"(12)�����ֽ��Ѽ�", 14, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,"Z"

	.Col = C_OWN_RGST_NO : .Row = -1 : .CellType = 4 : .TypePicMask = "999-99-99999" 
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_DEBIT_CREDIT_NM,C_DEBIT_CREDIT_NM,True)
	Call ggoSpread.SSSetColHidden(C_DEBIT_CREDIT,C_DEBIT_CREDIT,True)
	Call ggoSpread.SSSetColHidden(C_COMPANY_TYPE,C_COMPANY_TYPE,True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
				
	Call InitSpreadComboBox()
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With

	' 2�� �׸��� 
	With frm1.vspdData2
	
	ggoSpread.Source = frm1.vspdData2	
   'patch version
    ggoSpread.Spreadinit "V20041222_2",,parent.gAllowDragDropSpread    
    
	.ReDraw = false
	
	'����� 2�ٷ�    
    .ColHeaderRows = 2   
    
    .MaxCols = C_W4 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    Call AppendNumberPlace("7","2","0")

    ggoSpread.SSSetEdit		C_SEQ_NO,	"����", 10,,,15,1
    ggoSpread.SSSetEdit		C_MINOR_NM,	"�ڵ��", 10,,,100,1
    ggoSpread.SSSetEdit		C_MINOR_CD,	"�ڵ�", 10,,,10,1
	ggoSpread.SSSetEdit		C_W1,		"(1)���α���", 10,,,100,1
	ggoSpread.SSSetEdit		C_W2,		"(2)�����α���", 15,,,100,1  
    ggoSpread.SSSetCombo	C_W3_1,		"��"    , 10, 2
    ggoSpread.SSSetCombo	C_W3_2,		"��"    , 10, 2
    ggoSpread.SSSetEdit		C_W3_3,		" ", 3, 2,,1,1
    ggoSpread.SSSetCombo	C_W3_4,		"��"    , 10, 2
    ggoSpread.SSSetCombo	C_W3_5,		"��"    , 10, 2
	ggoSpread.SSSetFloat	C_W4,		"(4)�ͱݺһ�����(%)" ,15,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
     
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_MINOR_CD,True)
	
	' �׸��� ��� ��ħ ���� 
	ret = .AddCellSpan(C_SEQ_NO		, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_MINOR_NM	, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_MINOR_CD	, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W1			, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W2			, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W3			, -1000, 5, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W4			, -1000, 1, 2)	' SEQ_NO ��ħ 
	
     ' ù��° ��� ��� ���� 
	.Row = -1000
	.Col = C_W3
	.Text = "(3)������"
		
	' �ι�° ��� ��� ���� 
	'.Row = -999	
	'.Col = C_W3_1
	'.Text = "��(%)"
	'.Col = C_W3_2
	'.Text = "��"
	'.Col = C_W3_3	
	'.Text = " "
	'.Col = C_W3_4
	'.Text = "��(%)"
	'.Col = C_W3_5
	'.Text = "��"
	'
	.rowheight(-999) = 12	' ���� ������ 
					
	Call InitSpreadComboBox2()
	
	.ReDraw = true
	
    Call SetSpreadLock2 
    
    End With   
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx

	' ��/�뺯 
	IntRetCD1 = CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", " (MAJOR_CD='W1004') ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)  

	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_COMPANY_TYPE
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_COMPANY_TYPE_NM
	End If
		  
	iCodeArr = vbTab & lgF0
    iNameArr = vbTab & lgF1

End Sub

Sub InitSpreadComboBox2()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx, i, sVal, sVal2, sVal3
    
    ggoSpread.Source = frm1.vspddata2
    
    sVal = " " & vbTab
    For i = 0 to 100 Step 10
		sVal = sVal & CStr(i) & vbTab
    Next
    
    sVal2 = " " & vbTab & "�ʰ�" & vbTab & "�̻�" & vbTab 
    sVal3 = " " & vbTab & "����" & vbTab & "�̸�" & vbTab

	ggoSpread.SetCombo sVal, C_W3_1
	ggoSpread.SetCombo sVal, C_W3_4
	ggoSpread.SetCombo sVal2, C_W3_2
	ggoSpread.SetCombo sVal3, C_W3_5

End Sub

' ȯ�溯��1 %�� �޺��� �ε��� ������ 
Function ReadCombo1(pVal)
	If pVal = "" Then
		ReadCombo1 = 0
	Else
		ReadCombo1 = (UNICDbl(pVal) / 10) + 1	' �������� 
	End If
End Function

' ȯ�溯��2 �ʰ��� �޺��� �ε��������� 
Function ReadCombo2(pVal)
	Select Case pVal
		Case ">"
			ReadCombo2 = 1
		Case ">="
			ReadCombo2 = 2
		Case "<"
			ReadCombo2 = 2
		Case "<="
			ReadCombo2 = 1
		Case Else
			ReadCombo2 = 0
	End Select
End Function

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
	ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
	'ggoSpread.SSSetRequired C_DOC_DATE, -1, -1
	'ggoSpread.SSSetRequired C_DOC_AMT, -1, -1
	ggoSpread.SSSetRequired C_COMPANY_NM, -1, -1
	ggoSpread.SSSetRequired C_STOCK_RATE, -1, -1
	ggoSpread.SSSetRequired C_DOC_AMT, -1, -1
	'ggoSpread.SSSetRequired C_ACQUIRE_AMT, -1, -1
	ggoSpread.SSSetRequired C_COMPANY_TYPE_NM, -1, -1
	ggoSpread.SSSetRequired C_JUKSU, -1, -1
	ggoSpread.SSSetRequired C_OWN_RGST_NO, -1, -1
	ggoSpread.SSSetRequired C_CO_ADDR, -1, -1
	ggoSpread.SSSetRequired C_REPRE_NM, -1, -1
	ggoSpread.SSSetRequired C_STOCK_CNT, -1, -1

    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadLock2()
    With frm1

    .vspdData2.ReDraw = False
    
	ggoSpread.SpreadLock C_SEQ_NO, -1, C_SEQ_NO
	ggoSpread.SpreadLock C_W1, -1, C_W1, -1
	ggoSpread.SpreadLock C_W2, -1, C_W2, -1
	ggoSpread.SpreadLock C_W3_3, -1, C_W3_3, -1

    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
 
  	ggoSpread.SSSetProtected C_SEQ_NO, pvStartRow, pvEndRow
 	'ggoSpread.SSSetRequired C_DOC_DATE, pvStartRow, pvEndRow
 	ggoSpread.SSSetRequired C_DOC_AMT, pvStartRow, pvEndRow
 	ggoSpread.SSSetRequired C_COMPANY_NM, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_STOCK_RATE, pvStartRow, pvEndRow
	'ggoSpread.SSSetRequired C_ACQUIRE_AMT, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_COMPANY_TYPE_NM, pvStartRow, pvEndRow
	'ggoSpread.SSSetRequired C_HOLDING_TERM, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_JUKSU, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired C_OWN_RGST_NO, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_CO_ADDR, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_REPRE_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_STOCK_CNT, pvStartRow, pvEndRow
        
    .vspdData.ReDraw = True
    
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
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	lgCurrGrid = TYPE_1
	
End Sub

'============================================  ��ȸ���� �Լ�  ====================================

'============================== ���۷��� �Լ�  ========================================

Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg
	
	If gSelframeFlg = TAB2 Then Exit Function
	
	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If

    'Call ggoOper.ClearField(Document, "2")	
    ggoSpread.Source = frm1.vspdData
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

'====================================== �� �Լ� =========================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
	lgCurrGrid = TYPE_1	' �⺻ �׸��� 
End Function

Function ClickTab2()	
	Dim i, blnChange

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)
	gSelframeFlg = TAB2
	lgCurrGrid = TYPE_2
	
End Function

'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110110100101111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()

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


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_DEBIT_CREDIT
				.Col = Col
				intIndex = .Value
				.Col = C_DEBIT_CREDIT_NM
				.Value = intIndex	
			Case  C_DEBIT_CREDIT_NM
				.Col = Col
				intIndex = .Value
				.Col = C_DEBIT_CREDIT
				.Value = intIndex		
			Case C_COMPANY_TYPE
				.Col = Col
				intIndex = .Value
				.Col = C_COMPANY_TYPE_NM
				.Value = intIndex	
			Case C_COMPANY_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_COMPANY_TYPE
				.Value = intIndex	
		End Select
	End With
End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblAmt1, dblAmt2, dblSum
	With frm1.vspdData
	
    .Row = Row
    .Col = Col

    If .CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(.text) < CDbl(.TypeFloatMin) Then
         .text = .TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    Select Case Col
		Case C_ACQUIRE_AMT, C_HOLDING_TERM
			.Col = C_ACQUIRE_AMT	: dblAmt1 = UNICDbl(.value)
			.Col = C_HOLDING_TERM	: dblAmt2 = UNICDbl(.value)
			
			if dblAmt2 > 365 then
				Call DisplayMsgBox("970028", parent.VB_INFORMATION, "��⺸���Ⱓ�� 365��", "X")     
				'.value = 0                     
			End If
			
			dblSum = dblAmt1 * dblAmt2
			.Col = C_JUKSU			: .value = dblSum
		Case C_DOC_AMT
			.Row = Row : .Col = Col
			If UNICDbl(.value) < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, "�ݾ�", "X")     
				.value = 0                     
				Exit Sub
			End If
		
    End Select
    
	End With
End Sub

Function GetHead(Byval pCol)
	With frm1.vspdData
		.Col = pCol : .Row = 0	: GetHead = .Text
	End With
End Function

Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row
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

Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 

End Sub

'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )
    Frm1.vspdData2.Row = Row
    Frm1.vspdData2.Col = Col

    If Frm1.vspdData2.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData2.text) < CDbl(Frm1.vspdData2.TypeFloatMin) Then
         Frm1.vspdData2.text = Frm1.vspdData2.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

End Sub


Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData2.Row = Row
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2

End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
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

    Call SetToolbar("1100110000001111")

	Call ClickTab1()
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
    ggoSpread.Source = frm1.vspdData
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
    Dim blnChange, dblSum
    Dim i
    
    FncSave = False                                                         
    blnChange = False
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    for i = 1 to frm1.vspdData.MaxRows
    frm1.vspdData.col = C_HOLDING_TERM
    frm1.vspdData.row = i
	
		if UNICDbl(frm1.vspdData.value) > 365 then
					Call DisplayMsgBox("970028", parent.VB_INFORMATION, "��⺸���Ⱓ�� 365��", "X")     
					Exit Function                    
		End If
    
    next    
    
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If    

	ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		blnChange = True
    End If

	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If 
		
	If blnChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
	End If
	
	dblSum = FncSumSheet(frm1.vspdData, C_DOC_AMT, 1, frm1.vspdData.MaxRows, false, -1, -1, "V")
	
	If dblSum < 0 Then
		Call DisplayMsgBox("WC0013", parent.VB_INFORMATION, "(�ݾ�)", "X")                          
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
    Dim IntRetCD
    Dim imRow
    Dim iRow, iSeqNo

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim uCountID

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG
    
    If lgCurrGrid = TYPE_2 Then 
		Call InitGrid2
		Exit Function
	End If
	
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else    
        imRow = AskSpdSheetAddRowCount()
        
        If imRow = "" Then
            Exit Function
        End If
    
    End If
   
	With frm1	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		
		.vspdData.ReDraw = False
			
		' SEQ_NO �� �׸��忡 �ִ� ���� 
		iSeqNo = GetMaxSpreadVal(.vspdData , C_SEQ_NO)	' �ִ�SEQ_NO�� ���ؿ´�.
			
		ggoSpread.InsertRow ,imRow	' �׸��� �� �߰�(����� ��� ����)
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1	' �׸��� ���󺯰� 
		
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1	' �߰��� �׸����� SEQ_NO�� �����Ѵ�.
			.vspdData.Row = iRow
			.vspdData.Col = C_SEQ_NO
			.vspdData.Text = iSeqNo
			iSeqNo = iSeqNo + 1		' SEQ_NO�� �����Ѵ�.
		Next				
		.vspdData.ReDraw = True	

		''SetSpreadColor .vspdData.ActiveRow    
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

Function InitGrid2()    
	Dim i, iRow, iCol, iMaxRows, ret, sField, sFrom , sWhere, arrMinorNm, arrMinorCd, arrSeqNo, arrRef, arrW1_W2, sOldW1, sOldW2
	Dim soldMinorNm, soldMinorCd, iSpanRowW1, iSpanRowW2, iSpanCntW1, iSpanCntW2
	
	If frm1.vspdData2.MaxRows > 0 Then Exit Function
	
	sField	= "	A.MINOR_NM, B.MINOR_CD, B.SEQ_NO, B.REFERENCE"
	sFrom	= " B_MINOR A " & vbCrLf
	sFrom	= sFrom	& " 	INNER JOIN B_CONFIGURATION B WITH (NOLOCK) ON A.MAJOR_CD=B.MAJOR_CD AND A.MINOR_CD=B.MINOR_CD "
	sWhere	= " A.MAJOR_CD='W2003' "

	Call CommonQueryRs(sField, sFrom, sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		soldMinorCd	= "" : iSpanCntW1 = 0 : iSpanCntW2 = 0 : iSpanRowW1 = 0 : iSpanRowW2 = 0 : iRow = 0
		arrMinorNm	= Split(lgF0 , Chr(11))
		arrMinorCd	= Split(lgF1 , Chr(11))
		arrSeqNo	= Split(lgF2 , Chr(11))
		arrRef		= Split(lgF3 , Chr(11))
		
		iMaxRows = UBound(arrMinorNm)
		
		For i = 0 to iMaxRows -1
			
			ggoSpread.InsertRow , 1
			.Row = iRow + 1 : iRow = iRow + 1 : .Col = C_SEQ_NO : .Value = iRow
			
			.Col = C_MINOR_NM	: .Value = arrMinorNm(i)
			.Col = C_MINOR_CD	: .Value = arrMinorCd(i)
			
			arrW1_W2 = Split(arrMinorNm(i), "|")	' �ڵ���� | �� �и��Ѵ�.
			
			If sOldW1 <> arrW1_W2(0) Then	'C_W1 �� 
				iSpanCntW1 = 1 : iSpanRowW1 = .Row
			Else
				iSpanCntW1 = iSpanCntW1 + 1
				ret = .AddCellSpan(C_W1, iSpanRowW1, 1, iSpanCntW1)	' A1-A5 ��ħ 
			End If
			
			If sOldW2 <> arrW1_W2(1) Then	' C_W2
				iSpanCntW2 = 1	: iSpanRowW2 = .Row
			Else 
				iSpanCntW2 = iSpanCntW2 + 1
				ret = .AddCellSpan(C_W2, iSpanRowW2, 1, iSpanCntW2)	' A1-A5 ��ħ 
			End If		
			
			.Col = C_W1	: .Value = arrW1_W2(0)
			.Col = C_W2	: .Value = arrW1_W2(1)
				
			For iCol = 1 To 6	' �÷� ���� 
				.Col = C_W2 + iCol 
				Select Case iCol
					Case 1, 4
						.Value = ReadCombo1(arrRef(i)) : i = i + 1
					Case 2, 5
						.Value = ReadCombo2(arrRef(i)) : i = i + 1
					Case 6
						.Value = arrRef(i) ' For ������ i ���� �����Ѵ�.
					Case 3
						.Value = "~"
				End Select

				
			Next

			sOldW1	= arrW1_W2(0)
			sOldW2	= arrW1_W2(1)
		Next
		
		Call SetSpreadLock2
		
	End With

End Function

' -- �׸��� span
Function SetGridSpan()
	Dim soldMinorCd, iSpanCntW1, iSpanCntW2, iSpanRowW1, iSpanRowW2, iRow, i, sMinorNm, arrW1_W2, sOldW1, sOldW2, ret, iMaxRows
	
	With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		
		soldMinorCd	= "" : iSpanCntW1 = 0 : iSpanCntW2 = 0 : iSpanRowW1 = 0 : iSpanRowW2 = 0 : iRow = 0

		
		iMaxRows = .MaxRows
		
		For i = 1 to iMaxRows 
			
			.Row = i : .Col = C_MINOR_NM : sMinorNm = .Text
			
			arrW1_W2 = Split(sMinorNm, "|")	' �ڵ���� | �� �и��Ѵ�.
			
			If sOldW1 <> arrW1_W2(0) Then	'C_W1 �� 
				iSpanCntW1 = 1 : iSpanRowW1 = .Row
			Else
				iSpanCntW1 = iSpanCntW1 + 1
				ret = .AddCellSpan(C_W1, iSpanRowW1, 1, iSpanCntW1)	' A1-A5 ��ħ 
			End If
			
			If sOldW2 <> arrW1_W2(1) Then	' C_W2
				iSpanCntW2 = 1	: iSpanRowW2 = .Row
			Else 
				iSpanCntW2 = iSpanCntW2 + 1
				ret = .AddCellSpan(C_W2, iSpanRowW2, 1, iSpanCntW2)	' A1-A5 ��ħ 
			End If	

			.Col = C_W1	: .Value = arrW1_W2(0)
			.Col = C_W2	: .Value = arrW1_W2(1)
			
			sOldW1	= arrW1_W2(0)
			sOldW2	= arrW1_W2(1)
		Next
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
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
	
    ggoSpread.Source = frm1.vspdData	
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
    If frm1.vspdData.MaxRows > 0 Or frm1.vspdData2.MaxRows > 0 Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		Call SetGridSpan
		
		' �������� ���� : ���ߵǸ� ���ȴ�.
		Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

		'1 ����üũ : �׸��� �� 
		If wgConfirmFlg = "N" Then
			'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1

			'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
			Call SetToolbar("1111111100011111")										<%'��ư ���� ���� %>

		Else
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLock	1, -1, frm1.vspdData.MaxCols
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLock	1, -1, frm1.vspdData2.MaxCols
			
			Call SetToolbar("1110000000011111")										<%'��ư ���� ���� %>
		End If
	Else
		Call SetToolbar("1110110100001111")										<%'��ư ���� ���� %>
	End If
	'frm1.vspdData.focus			
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
    Dim strVal, strDel, lMaxRows, lMaxCols
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With Frm1
	
		With frm1.vspdData
		
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
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
       .txtSpread.value        =  strDel & strVal
       strDel = ""	: strVal = ""
       
       ' 2�� �׸��� 
		With frm1.vspdData2
			
			ggoSpread.Source = frm1.vspdData2
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
       
       frm1.txtSpread2.value        =  strDel & strVal
		.txtMode.value        =  Parent.UID_M0002
		'.txtUpdtUserId.value  =  Parent.gUsrID
		'.txtInsrtUserId.value =  Parent.gUsrID
		.txtCurrGrid.value     = lgCurrGrid
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" width=200>
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ͱݺһ����� ���</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:GetRef">�ݾ׺ҷ�����</A>&nbsp;</TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w2103ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
								<TD HEIGHT="100%">
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
									<script language =javascript src='./js/w2103ma1_vaSpread1_vspdData.js'></script>
								</DIV>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
									<script language =javascript src='./js/w2103ma1_vaSpread2_vspdData2.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

