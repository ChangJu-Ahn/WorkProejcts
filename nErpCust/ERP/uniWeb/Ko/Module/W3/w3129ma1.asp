
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

Const BIZ_MNU_ID = "w3129ma1"
Const BIZ_PGM_ID = "w3129mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "w3129OA1"

Dim C_W1
Dim C_W1_NM
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6

Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
    C_W1				= 1
    C_W1_NM				= 2
    C_W2				= 3
    C_W3				= 4
    C_W4				= 5
    C_W5				= 6
    C_W6				= 7
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

	'����� 2�ٷ�    
    .ColHeaderRows = 2   
    
    .MaxCols = C_W6 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
    ggoSpread.SSSetEdit		C_W1,		"�ڵ�", 10,,,10,1
	ggoSpread.SSSetEdit		C_W1_NM,	"(1)�ڻ� ����", 25,,,100,1
	ggoSpread.SSSetFloat	C_W2,		"(2)�հ�ױݾ�", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  
    ggoSpread.SSSetFloat	C_W3,		"(3)���๰", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  
    ggoSpread.SSSetFloat	C_W4,		"(4)�����ġ", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  
    ggoSpread.SSSetFloat	C_W5,		"(5)��Ÿ�ڻ�", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  
    ggoSpread.SSSetFloat	C_W6,		"(6)���������ڻ�", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec  

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)

	' �׸��� ��� ��ħ ���� 
	ret = .AddCellSpan(C_W1		, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W1_NM	, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W2		, -1000, 1, 2)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W3		, -1000, 3, 1)	' SEQ_NO ��ħ 
	ret = .AddCellSpan(C_W6		, -1000, 1, 2)	' SEQ_NO ��ħ 
	
     ' ù��° ��� ��� ���� 
	.Row = -1000
	.Col = C_W3
	.Text = "�� �� �� �� �� ��"
		
	' �ι�° ��� ��� ���� 
	.Row = -999	
	.Col = C_W3
	.Text = "(3)���๰"
	.Col = C_W4
	.Text = "(4)�����ġ"
	.Col = C_W5	
	.Text = "(5)��Ÿ�ڻ�"
	
	.rowheight(-999) = 15	' ���� ������ 
					
	'Call InitSpreadComboBox2()
	
	.ReDraw = true

	Call SetSpreadLock 

    End With   
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_W1, -1, C_W2
	
	ggoSpread.SSSetProtected C_W3, 3, 3
	ggoSpread.SSSetProtected C_W4, 3, 3 
	ggoSpread.SSSetProtected C_W5, 3, 3 
	ggoSpread.SSSetProtected C_W6, 3, 3
	
	ggoSpread.SSSetProtected C_W3, 6, 7
	ggoSpread.SSSetProtected C_W4, 6, 7 
	ggoSpread.SSSetProtected C_W5, 6, 7 
	ggoSpread.SSSetProtected C_W6, 6, 7
	
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
	
	
End Sub

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

	iMaxRows = 7 ' �ϵ��ڵ��Ǵ� ��� 
	With frm1.vspdData
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow , iMaxRows

		iRow = 0
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1		: .value = iRow

		.Redraw = True
		
		Call InitData2
		
		Call SetSpreadLock
	End With	
End Sub

 ' -- DBQueryOk ������ �ҷ��ش�.
Sub InitData2()
	Dim iRow
	
	With frm1.vspdData
		.Redraw = False

		iRow = 0
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1_NM	: .value = " �� �� | (101)�⸻ �����"
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1_NM	: .value = " �� �� | (102)�����󰢴����"
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1_NM	: .value = " ǥ �� | (103)�̻��ܾ�"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1_NM	: .value = "(104)�� ������"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1_NM	: .value = "(105)ȸ��ձݰ���"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1_NM	: .value = " �� �� | (106)�󰢺��ξ�"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_W1_NM	: .value = " �� �� | (107)���κ��ξ�"
	End With
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

    Dim arrW3, arrW4, arrW5, arrW6, iRow
	call CommonQueryRs("W3, W4, W5, W6","dbo.ufn_TB_20_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 = "" Then	 Exit Function
	
	arrW3 = Split(lgF0, chr(11))
	arrW4 = Split(lgF1, chr(11))
	arrW5 = Split(lgF2, chr(11))
	arrW6 = Split(lgF3, chr(11))
	
	With frm1.vspdData
		.Redraw = False
		
		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData
		
		For iRow = 1 To 3
			.Row = iRow

			If lgIntFlgMode = parent.OPMD_UMODE Then
				ggoSpread.UpdateRow iRow
			End If
			
			.Col = C_W3 : .Value = arrW3(iRow-1)
			.Col = C_W4 : .Value = arrW4(iRow-1)
			.Col = C_W5 : .Value = arrW5(iRow-1)
			.Col = C_W6 : .Value = arrW6(iRow-1)
			
			' �հ���� ����Ѵ�.
			Call FncSumSheet(frm1.vspdData, iRow, C_W3, C_W6, true, iRow, C_W2, "H")
		Next
			
		.Redraw = True
	End With
	
End Function

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1110100000000111")										<%'��ư ���� ���� %>

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


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
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

	Dim dblSum, dblCol(1)
	Dim StrC4W3,StrC4W4,StrC4W5,StrC4W6
	Dim StrC5W3,StrC5W4,StrC5W5,StrC5W6
	
	With frm1.vspdData
	
		Select Case Col
			Case C_W3, C_W4, C_W5, C_W6
				.Col = Col	: .Row = Row	: dblSum = UNiCDbl(.Value)
				If dblSum < 0 Then
					Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, dblSum, "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
					.Value = 0
				End If
				
				' �հ���� ����Ѵ�.
				Call FncSumSheet(frm1.vspdData, Row, C_W3, C_W6, true, Row, C_W2, "H")
				
				' Row = 1, Row = 2
				If Row = 1 Or Row = 2 Then
					.Col = Col
					.Row = 1 : dblCol(0) = UNICDbl(.value)
					.Row = 2 : dblCol(1) = UNICDbl(.value)
					dblSum = dblCol(0) - dblCol(1)
					.Row = 3 : .value = dblSum
					
					' �հ���� ����Ѵ�.
					Call FncSumSheet(frm1.vspdData, 3, C_W3, C_W6, true, 3, C_W2, "H")
					ggoSpread.UpdateRow 3
				End If
				
				if Row = 4 Or Row = 5 Then
				  .Col = Col
				  Select case col
				  
					Case C_W3
						.Col = C_W3
						.Row = 4 : StrC4W3 = UNICDbl(.value)
						.Row = 5 : StrC5W3 = UNICDbl(.value)

						If	StrC4W3 > StrC5W3 then
							dblSum = StrC4W3 - StrC5W3
							.Row = 6 : .value = 0
							.Row = 7 : .value = dblSum
						Else
						    dblSum = StrC5W3 - StrC4W3
						    .Row = 6 : .value = dblSum
						    .Row = 7 : .value = 0
						End if

					Case C_W4
						.Col = C_W4
						.Row = 4 : StrC4W4 = UNICDbl(.value)
						.Row = 5 : StrC5W4 = UNICDbl(.value)
						If	StrC4W4 > StrC5W4 then
							dblSum = StrC4W4 - StrC5W4
							.Row = 6 : .value = 0
							.Row = 7 : .value = dblSum
						Else
						    dblSum = StrC5W4 - StrC4W4
						    .Row = 6 : .value = dblSum
						    .Row = 7 : .value = 0
						End if

					Case C_W5
						.Col = C_W5
						.Row = 4 : StrC4W5 = UNICDbl(.value)
						.Row = 5 : StrC5W5 = UNICDbl(.value)
						If	StrC4W5 > StrC5W5 then
							dblSum = StrC4W5 - StrC5W5
							.Row = 6 : .value = 0
							.Row = 7 : .value = dblSum
						Else
						    dblSum = StrC5W5 - StrC4W5
						    .Row = 6 : .value = dblSum
						    .Row = 7 : .value = 0
						End if

					Case C_W6
						.Col = C_W6
						.Row = 4 : StrC4W6 = UNICDbl(.value)
						.Row = 5 : StrC5W6 = UNICDbl(.value)
						If	StrC4W6 > StrC5W6 then
							dblSum = StrC4W6 - StrC5W6
							.Row = 6 : .value = 0
							.Row = 7 : .value = dblSum
						Else
						    dblSum = StrC5W6 - StrC4W6
						    .Row = 6 : .value = dblSum
						    .Row = 7 : .value = 0
						End if
						ggoSpread.UpdateRow 4
						ggoSpread.UpdateRow 5
				  End Select
				 
				 Call FncSumSheet(frm1.vspdData, 6, C_W3, C_W6, true, 6, C_W2, "H")
				 Call FncSumSheet(frm1.vspdData, 7, C_W3, C_W6, true, 7, C_W2, "H")
				 ggoSpread.UpdateRow 6
				 ggoSpread.UpdateRow 7
				 'Call W2SumData()
				 
				 End if 
					
									
		End Select
	
	End With
End Sub


Sub W2SumData()

	Dim StrC4W2,StrC5W2,StrSum
    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.Col = C_W2
    frm1.vspdData.Row = 4
    StrC4W2 = UNICDbl(frm1.vspddata.value)
    frm1.vspdData.Row = 5
    StrC5W2 = UNICDbl(frm1.vspddata.value)
    
    If	StrC4W2 > StrC5W2 then
		StrSum = StrC4W2 - StrC5W2
		frm1.vspdData.Row = 6
		frm1.vspdData.value = 0
		frm1.vspdData.Row = 7
		frm1.vspdData.value = StrSum
	Else
		StrSum = StrC5W2 - StrC4W2
		frm1.vspdData.Row = 6
		frm1.vspdData.value = StrSum
		frm1.vspdData.Row = 7
		frm1.vspdData.value = 0		
    End if
    
    ggoSpread.UpdateRow 6
	ggoSpread.UpdateRow 7
    
End Sub

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
    Dim blnChange, dblSum, iCol
    
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

	If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
	      Exit Function
	End If    
	
	If blnChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
	End If
	
	With frm1.vspdData
	
		For iCol = C_W3 To C_W6
			.Row = 3 : .Col = iCol
			If UNICDbl(.value) < 0 Then
				Call DisplayMsgBox("WC0006", parent.VB_INFORMATION, GetColName(iCol), "X")           '��: "%1 �ݾ��� 0���� �����ϴ�."
				Exit Function
			End If
		Next

	End With	

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
	Call InitData2 ' �׸��� �ڻ걸�� ��� 
	
	lgIntFlgMode = parent.OPMD_UMODE
		    
	' �������� ���� : ���ߵǸ� ���ȴ�.
	Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	'1 ����üũ : �׸��� �� 
	If wgConfirmFlg = "N" Then
		'ggoSpread.SpreadUnLock	C_W1, -1, C_W5, -1
		Call SetSpreadLock()

		'2 ���ȯ�氪 , �ε��ȯ�氪 �� 
		Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>

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
    
	With frm1.vspdData
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
				For lCol = C_W1 To lMaxCols
					.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
				Next
				strVal = strVal & Trim(.Text) &  Parent.gRowSep
		  End If  
		Next
       
        frm1.txtSpread.value        =  strDel & strVal
		frm1.txtMode.value        =  Parent.UID_M0002

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
End Function


Function DbSaveOk()													        <%' ���� ������ ���� ���� %>
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
					<TD WIDTH=* align=right><A href="vbscript:GetRef()">�ݾ׺ҷ�����</A>| <A href="vbscript:OpenRefMenu">�ҵ�ݾ��հ�ǥ��ȸ</A></TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w3129ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
									<script language =javascript src='./js/w3129ma1_vaSpread1_vspdData.js'></script>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

