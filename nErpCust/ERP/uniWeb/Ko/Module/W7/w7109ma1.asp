<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ÿ�������� 
'*  3. Program ID           : w7109mA1
'*  4. Program Name         : w7109mA1.asp
'*  5. Program Desc         : ��50ȣ(��) �ں��ݰ� ��������������(��)
'*  6. Modified date(First) : 2005/02/21
'*  7. Modified date(Last)  : 2005/02/21
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
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "w7109mA1"
Const BIZ_PGM_ID		= "w7109mB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_REF_PGM_ID	= "w7109mB2.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		 = "w7109OA1"
Const JUMP_PGM_ID		= "W5101MA1"

' -- �׸��� �÷� ���� 
Dim C_SEQ_NO
Dim C_W1
Dim C_W1_BT
Dim C_W1_NM
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
DIm C_W_DESC

Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 
Dim lgFISC_START_DT, lgFISC_END_DT 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO	= 1
	C_W1		= 2
	C_W1_BT		= 3
	C_W1_NM		= 4
	C_W2		= 5
	C_W3		= 6
	C_W4		= 7
	C_W5		= 8
	C_W_DESC	= 9
	
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



'============================================  �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
End Sub


Sub InitSpreadComboBox()
    Dim IntRetCD1

End Sub

Function OpenAdItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�������� �˾�"					<%' �˾� ��Ī %>
	arrParam(1) = "TB_ADJUST_ITEM"					<%' TABLE ��Ī %>
	

		frm1.vspdData.Col = C_W1
	    arrParam(2) = frm1.vspdData.Text		<%' Code Condition%>

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
		Call SetAdItem(arrRet)
	End If	
	
End Function

Function SetAdItem(byval arrRet)
    With frm1
			.vspdData.Col = C_W1
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_W1_NM
			.vspdData.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
	End With
End Function

Function GetAdItem(ByVal pCol, ByVal pRow)
	Dim arrRet(2), sWhere, bRet

	If pCol = C_W1 Then
		sWhere = " ITEM_CD LIKE '%"
	ElseIf pCol = C_W1_NM Then
		sWhere = " ITEM_NM LIKE '%"
	Else
		Exit Function
	End If
	
	With frm1.vspdData
		.Col = pCol
		If .Text <> "" Then
			sWhere = sWhere & .Text & "%' "		<%' Code Condition%>
		
			bRet = CommonQueryRs("top 1 ITEM_CD,ITEM_NM"," TB_ADJUST_ITEM ",sWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			arrRet(0) = Replace(lgF0, chr(11), "")
			arrRet(1) = Replace(lgF1, chr(11), "")
		Else
			arrRet(0) = ""
			arrRet(1) = ""
		End If
	End With
	
	Call SetAdItem(arrRet)
	
End Function


Sub InitSpreadSheet()
	Dim ret, iRow
	
    Call initSpreadPosVariables()  

	Call AppendNumberPlace("6","3","2")
 	
	' 1�� �׸��� 

	With Frm1.vspdData
				
		ggoSpread.Source = Frm1.vspdData
		'patch version
		ggoSpread.Spreadinit "V20041222_0" ,,parent.gAllowDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W_DESC + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		'����� 2�ٷ�    
	    .ColHeaderRows = 2

	    ggoSpread.SSSetEdit		C_SEQ_NO,	"����",				5,,,6,1	' �����÷� 
		ggoSpread.SSSetEdit		C_W1,		"(1)���� �Ǵ� ����",			7,,,10,1
	    ggoSpread.SSSetButton 	C_W1_BT
		ggoSpread.SSSetEdit		C_W1_NM,	"(1)�����",	15,,,50,1
		ggoSpread.SSSetFloat	C_W2,		"(2)�����ܾ�",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	    ggoSpread.SSSetFloat	C_W3,		"(3)����",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
	    ggoSpread.SSSetFloat	C_W4,		"(4)����",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
		ggoSpread.SSSetFloat	C_W5,		"(5)�⸻�ܾ�" & vbCrLf & "(�ͱ�������)",			15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"","" 
		ggoSpread.SSSetEdit		C_W_DESC,	"���",	20,,,100,1

	    ret = .AddCellSpan(0, -1000, 1, 2)
	    ret = .AddCellSpan(C_SEQ_NO, -1000, 1, 2)
	    ret = .AddCellSpan(C_W1, -1000, 3, 2)
	    ret = .AddCellSpan(C_W2, -1000, 1, 2)
	    ret = .AddCellSpan(C_W3, -1000, 2, 1)
	    ret = .AddCellSpan(C_W5, -1000, 1, 2)
	    ret = .AddCellSpan(C_W_DESC, -1000, 1, 2)
	    
	    ' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_W1
		.Text = "(1)���� �Ǵ� ����"
		.Col = C_W3
		.Text = "���������"
	
		' �ι�° ��� ��� ���� 
		.Row = -999
		.Col = C_W3
		.Text = "(3)����"
		.Col = C_W4
		.Text = "(4)����"

		.rowheight(-999) = 15	' ���� ������ 
		
		Call ggoSpread.MakePairsColumn(C_W1,C_W1_BT)

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO,C_SEQ_NO,True)
		
		Call SetSpreadLock()

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


Sub SetSpreadLock()

	With Frm1.vspdData
	
		ggoSpread.Source = Frm1.vspdData

		ggoSpread.SpreadUnLock C_W1, -1, C_W_DESC	' ��ü ���� 

	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim iRow

	With Frm1.vspdData
		ggoSpread.Source = Frm1.vspdData

		ggoSpread.SpreadUnLock C_W1, pvStartRow, C_W_DESC, pvEndRow	' ��ü ���� 
		ggoSpread.SSSetRequired C_W1, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired C_W1_NM, pvStartRow, pvEndRow
		ggoSpread.SpreadLock C_W5,   -1, C_W5

		If .MaxRows <= pvEndRow And .MaxRows >= pvStartRow Then
			ggoSpread.SpreadLock C_W1,   .MaxRows, C_W5, .MaxRows
		End If
			
	End With	
End Sub

Sub SetSpreadTotalLine()
	Dim ret
		
	ggoSpread.Source = Frm1.vspdData
	With Frm1.vspdData
		If .MaxRows > 0 Then
			.Row = .MaxRows
			ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' ���� 2�� ��ħ 
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "��"	:	.TypeHAlign = 2
			SetSpreadColor 1, .MaxRows

		End If
	End With
End Sub 

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_SEQ_NO	= iCurColumnPos(1)
            C_W1		= iCurColumnPos(2)
            C_W1_BT		= iCurColumnPos(3)
            C_W1_NM		= iCurColumnPos(4)
            C_W2		= iCurColumnPos(5)
            C_W3		= iCurColumnPos(6)
            C_W4		= iCurColumnPos(7)
            C_W5		= iCurColumnPos(8)
            C_W_DESC	= iCurColumnPos(9)
    End Select    
End Sub

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg
     wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
	 
	ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
			
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_REF_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key           
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With  
	
End Function

Function GetRefOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr, iSeqNo, iLastRow, iRow
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
	    lgIntFlgMode = parent.OPMD_CMODE
		

		Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>
		Call SetSpreadColor(1, Frm1.vspdData.MaxRows)
		Call SetSpreadTotalLine
'		Call Fn_GridCalc(0, 0)
		Call ChangeRowFlg(frm1.vspdData)
		lgBlnFlgChgValue = True
	End If

	frm1.vspdData.focus			
End Function

Function ChangeRowFlg(iObj)
	Dim iRow
	
	With iObj
		ggoSpread.Source = iObj
		
		For iRow = 1 To .MaxRows
			.Col = 0 : .Row = iRow : .Value = ggoSpread.InsertFlag
			Call Fn_GridCalc(C_W2, iRow)
		Next
	End With
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

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100111100000111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	
 
	Call InitComboBox	' �����ؾ� �Ѵ�. ����� ȸ��������� �о���� ���� 
	Call InitData

'	Call DBQuery()
	Call FncQuery
	
     

End Sub

'============================================  ����� �Լ�  ====================================
Function Fn_GridCalc(ByVal pCol, ByVal pRow)
	Dim iRow, dblSum
	Dim dblW2, dblW3, dblW4, dblW5
	
	If Frm1.vspdData.MaxRows <= 0 Then Exit Function

    ggoSpread.Source = Frm1.vspdData
    iRow = pRow

	With Frm1.vspdData
		' (5) = (2) - (3) + (4)
		If pRow = 0 Then iRow = Frm1.vspdData.ActiveRow
		.Row = iRow	:	.Col = C_W2	:	dblW2 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W3	:	dblW3 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W4	:	dblW4 = UNICdbl(.Text)
		.Row = iRow	:	.Col = C_W5	:	dblW5 = dblW2 - dblW3 + dblW4
		If dblW5 = 0 And dblW2 = 0 And dblW3 = 0 And dblW4 = 0 Then
			.Text = ""
		Else
			.Text = dblW5
		End If
			
	End With

	dblSum = FncSumSheet(Frm1.vspdData, C_W2, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W2, "V")	' �հ� 
	dblSum = FncSumSheet(Frm1.vspdData, C_W3, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W3, "V")	' �հ� 
	dblSum = FncSumSheet(Frm1.vspdData, C_W4, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W4, "V")	' �հ� 
	dblSum = FncSumSheet(Frm1.vspdData, C_W5, 1, Frm1.vspdData.MaxRows - 1, true, Frm1.vspdData.MaxRows, C_W5, "V")	' �հ� 
    
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.UpdateRow Frm1.vspdData.MaxRows
End Function


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

'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim iIdx, iRow, sW3, sW4, dblW2

	With Frm1.vspdData
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
		
		

	End With
End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim dblSum
	Dim arrVal
	
	lgBlnFlgChgValue= True ' ���濩�� 
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.UpdateRow Row

	With frm1.vspdData

	If Col = C_W1 Then
		If Len(.Text) > 0 Then
			.Row = Row

			.Col = Col
			If CommonQueryRs("ITEM_NM", " TB_ADJUST_ITEM (NOLOCK)" , "ITEM_CD = '" & .Text &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		    	.Col	= C_W1_NM
		    	arrVal				= Split(lgF0, Chr(11))
				.Text	= arrVal(0)
			Else
				.Text	= ""
				.Col	= C_W1_NM
				.Text	= ""
			End If
		Else
			.Col = C_W1_NM
			.Text = ""
		End If
	ElseIf Col = C_W1_NM Then
		If Len(.Text) > 0 Then
			.Row = Row

			.Col = Col
			If CommonQueryRs("ITEM_CD", " TB_ADJUST_ITEM (NOLOCK)" , "ITEM_NM = '" & .Text &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		    	.Col	= C_W1
		    	arrVal				= Split(lgF0, Chr(11))
				.Text	= arrVal(0)
			Else
				.Text	= ""
				.Col	= C_W1
				.Text	= ""
			End If
		Else
			.Col = C_W1
			.Text = ""
		End If
	End If

	End With

	' --- �߰��� �κ� 
	Call Fn_GridCalc(Col, Row)

'	Call GetAdItem(Col, Row)			' �������� ��������		ǥ�ؿ��� ��÷� ��Ī�� �Ȱ����´�.

	
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = Frm1.vspdData
   
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
    	Exit Sub
       ggoSpread.Source = Frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	Frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = Frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If Frm1.vspdData.MaxRows = 0 Then
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

	ggoSpread.Source = Frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = Frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
'    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With Frm1.vspdData
		If Row > 0 And Col = C_W1_BT Then
		    .Row = Row
		    .Col = C_W1_BT

		    Call OpenAdItem()
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
    If lgBlnFlgChgValue = False Then
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

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

	With Frm1.vspdData
	    ggoSpread.Source = Frm1.vspdData
	    iActiveRow = .ActiveRow

		If .ActiveRow > 0 and .ActiveRow <> .MaxRows Then
			.focus
			.ReDraw = False
		
			ggoSpread.CopyRow
			.Col = C_W1
			.Text = ""

			Call SetDefaultVal(iActiveRow + 1, 1)
			SetSpreadColor iActiveRow, iActiveRow + 1
			.ReDraw = True
			
			Call Fn_GridCalc(C_W2, iActiveRow + 1)
    
		End If
	End With


    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
    Dim lDelRows, dblSum 

		With Frm1.vspdData
			.focus

			ggoSpread.Source = Frm1.vspdData
			If .MaxRows <= 0 Then
				Exit Function
			ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
				MsgBox "�հ�� ������ �� �����ϴ�.", vbCritical
				Exit Function
			Else
				lDelRows = ggoSpread.EditUndo
				lgBlnFlgChgValue = True
				lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
				If lDelRows > 0 Then ggoSpread.EditUndo lDelRows
			End If
			
		End With

		Call Fn_GridCalc(0, 0)


End Function


' -- �հ� ������ üũ(Header Grid)
Function CheckTotRow(Byref pObj, Byval pRow) 
	CheckTotRow = False
	pObj.Col = C_SEQ_NO : pObj.Row = pRow
	If pObj.Text = "" Then Exit Function
	If pObj.Text = "999999" And pObj.MaxRows > 1 Then	 ' �հ� �� 
		CheckTotRow = True
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

	With Frm1.vspdData
	
		.focus
		ggoSpread.Source = Frm1.vspdData
	
		iSeqNo = .MaxRows+1
	
		if .MaxRows = 0 then
		
			ggoSpread.InsertRow  imRow 
			.Col	= C_SEQ_NO	:	.Text	= 1
			
			ggoSpread.InsertRow  imRow 
			.Row = .MaxRows
			.Col	= C_SEQ_NO	:	.Text	= 999999
			ret = .AddCellSpan(C_W1	, .MaxRows, 3, 1)	' ���� 2�� ��ħ 
			.Col	= C_W1	:	.CellType = 1	:	.Text	= "��"	:	.TypeHAlign = 2
			SetSpreadColor 1, .MaxRows
			.Row  = 1
			.ActiveRow = 1

		else
			iRow = .ActiveRow

			If iRow = .MaxRows Then	' -- ������ �հ��ٿ��� InsertRow�� �ϸ� ������ �߰��Ѵ�.
				iRow = iRow - 1
				ggoSpread.InsertRow iRow, imRow 
				SetSpreadColor iRow, iRow + imRow

				Call SetDefaultVal(iRow + 1, imRow)
			Else
				ggoSpread.InsertRow ,imRow
				SetSpreadColor iRow + 1, iRow + imRow

				Call SetDefaultVal(iRow + 1, imRow)
			End If   
			.vspdData.Row  = iRow + 1
			.vspdData.ActiveRow = iRow +1
			
        End if 	
		
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


' �׸��忡 SEQ_NO, TYPE �ִ� ���� 
Function SetDefaultVal(iRow, iAddRows)
	
	Dim i, iSeqNo
	
	With Frm1.vspdData
	
		If iAddRows = 1 Then ' 1�ٸ� �ִ°�� 
			.Row = iRow
			.Value = MaxSpreadVal(Frm1.vspdData, C_SEQ_NO, iRow)
		Else
			iSeqNo = MaxSpreadVal(Frm1.vspdData, C_SEQ_NO, iRow)	' ������ �ִ�SeqNo�� ���Ѵ� 
			
			For i = iRow to iRow + iAddRows -1
				.Row = i	:	.Col = C_SEQ_NO
				If .Text <> 999999 Then
					: .Value = iSeqNo : iSeqNo = iSeqNo + 1
				End If
			Next
		End If
	End With
End Function

Function FncDeleteRow() 
    Dim lDelRows, iActiveRow, dblSum 

	With Frm1.vspdData
		.focus

		ggoSpread.Source = Frm1.vspdData
		If .MaxRows <= 0 Then
			Exit Function
		ElseIf CheckTotRow(Frm1.vspdData, .ActiveRow) = True Then
			MsgBox "�հ�� ������ �� �����ϴ�.", vbCritical
			Exit Function
		Else
			lDelRows = ggoSpread.DeleteRow
			lgBlnFlgChgValue = True
			lDelRows = CheckLastRow(Frm1.vspdData, lDelRows)
			If lDelRows > 0 Then ggoSpread.DeleteRow lDelRows
		End If
		
	End With

	Call Fn_GridCalc(0, 0)
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
	
    ggoSpread.Source = Frm1.vspdData
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
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	If Frm1.vspdData.MaxRows > 0  Then
		'-----------------------
		'Reset variables area
		'-----------------------
		lgIntFlgMode = parent.OPMD_UMODE
		

		Call SetToolbar("110111110000111")										<%'��ư ���� ���� %>
	End If
	
	Call SetSpreadTotalLine ' - �հ���� �籸�� 
	
	Frm1.vspdData.focus			
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
    
	With Frm1.vspdData

		ggoSpread.Source = Frm1.vspdData
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

	Frm1.txtSpread.value      = strVal
	strVal = ""

	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow											        <%' ���� ������ ���� ���� %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:GetRef">�ݾ׺ҷ�����</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w7109ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%> CLASS="TD61" BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;�������������ҵ��� 
										</TD>
									</TR>
									<TR>
										<TD >
											<script language =javascript src='./js/w7109ma1_vspdData_vspdData.js'></script>
										</TD>
									</TR>
									
								</TABLE>
							    </TD>
							</TR>
						</TABLE>
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
							<A href="Vbscript:ProgramJump()">��15ȣ ����������</A>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>