
<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% session.CodePage=949 %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ����� ��������(��)
'*  3. Program ID           : W3113MA1
'*  4. Program Name         : W3113MA1.asp
'*  5. Program Desc         : ����� ��������(��)
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2005/03/18
'*  8. Modifier (First)     : ȫ���� 
'*  9. Modifier (Last)      : ȫ���� 
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

Const BIZ_MNU_ID = "W3113MA1"
Const BIZ_PGM_ID = "W3113MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID = "W3113OA1"

dim C_H1
Dim C_ACCT_NM
Dim C_COL1
Dim C_COL2
Dim C_COL3
Dim C_COL4
Dim C_SUM
Dim C_Row1 	
Dim C_Row2 	
Dim C_Row3 	
Dim C_Row4 	
Dim C_Row5 	
Dim C_Row6 	
Dim C_Row7 	
Dim C_Row8 	
Dim C_Row9 	
Dim C_Row10	
Dim C_Row11	
Dim C_Row12
dim C_Row13
dim C_Row14


Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_H1		=	1
	C_ACCT_NM		=	2
	C_COL1			=	3
	C_COL2			=	4
	C_COL3			=	5
	C_COL4			=	6
	C_SUM			=	7
	

	
	C_Row1 		=	1
	C_Row2 		=	2
	C_Row3 		=	3
	C_Row4 		=	4
	C_Row5 		=	5
	C_Row6 		=	6
	C_Row7 		=	7
	C_Row8 		=	8
	C_Row9 		=	9
	C_Row10		=	10
	C_Row11		=	11
	C_Row12		=	12
	C_Row13		=	13
	C_Row14		=	14

	
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


'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '


    
 

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
	
    .MaxCols = C_SUM + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
    
  
    ggoSpread.SSSetEdit		C_ACCT_NM,		"", 20,,,45,1
	ggoSpread.SSSetFloat	C_COL1,		"1", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec 
	ggoSpread.SSSetFloat	C_COL2,		"2", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_COL3,		"3", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_COL4,		"4", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
    ggoSpread.SSSetFloat	C_SUM,		"�հ�", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec 

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		  
	.ReDraw = true

	Call SetSpreadLock 

    End With   
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()

End Sub

Sub SetSpreadLock()
	DIM ret
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_h1, -1, C_h1
    ggoSpread.SpreadLock C_ACCT_NM, -1, C_ACCT_NM
    ggoSpread.SpreadLock C_SUM    , -1, C_SUM
    ggoSpread.SpreadLock C_COL1, C_Row12, C_COL4 ,C_Row12
	ggoSpread.SpreadLock C_COL1, C_Row14, C_COL4 ,C_Row14
    
  
    .vspdData.ReDraw = True

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
	Dim iMaxRows, iRow
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	iMaxRows = c_row14

	With frm1.vspdData
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow , iMaxRows
		.Redraw = True
	
		Call InitData2
		Call SetSpreadLock
		
	End With	
End Sub

 ' -- DBQueryOk ������ �ҷ��ش�.
Sub InitData2()
	Dim iRow
	on error resume next
	With frm1.vspdData
		.Redraw = False

		iRow = 0
		iRow = iRow + 1 : .Row = iRow
		.Col = C_H1	: .value = " (6)��������"
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_H1	: .value = " (7)�����ݾ�"
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_H1	: .value = " (8)���������߻��������"
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_H1	: .value = " (9)����� �ش�ݾ�"

		iRow = iRow + 1 : .Row = iRow: .TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Col = C_H1	:.value = "(10)"& vbCrLf & "�ſ�" & vbCrLf & "ī���" & vbCrLf & "���" & vbCrLf & "�ݾ�"

		.Col = C_H1+1	:	.TypeHAlign = 3	:	.TypeVAlign = 2
		.Col = C_ACCT_NM	: .value = "(11)�������10�����ʰ���"
		iRow = iRow + 1 : .Row = iRow: .TypeEditMultiLine = True	:	.TypeHAlign = 3	:	.TypeVAlign = 2
		
		iRow = iRow + 1 : .Row = iRow : .TypeEditMultiLine = True	:	.TypeHAlign = 3 :	.TypeVAlign = 2
		.Col = C_ACCT_NM	: .value = "(12)�������������" & vbCrLf & "(��41����1��)"
		'.rowheight(iRow) = 20	
		iRow = iRow + 1 
		
		iRow = iRow + 1 : .Row = iRow: .TypeEditMultiLine = True	:	.TypeHAlign = 3 :	.TypeVAlign = 2
		.Col = C_ACCT_NM	: .value = "(13)�������5�����ʰ���" & vbCrLf & "((11)��(12)����)"
		'.rowheight(iRow) = 20	
		iRow = iRow + 1 
		
		iRow = iRow + 1 : .Row = iRow
		.Col = C_ACCT_NM	: .value = "(14)��    Ÿ"

		iRow = iRow + 1 : .Row = iRow
		.Col = C_ACCT_NM	: .value = "(15)�ſ�ī��� �����հ�"
		
		iRow = iRow + 1 : .Row = iRow: .TypeEditMultiLine = True	:	.TypeHAlign = 2	:	.TypeVAlign = 2
		.Col = C_H1	: .value = " (16) �ſ�ī��� �̻�� ���ξ�"
		iRow = iRow + 1 : .Row = iRow
		.Col = C_H1	: .value = " (15)�������ξ�((8)+(16))"

		.Row = 1
		.Col = C_COL1	: .CellType = 1	: .TypeHAlign = 2
		.Row = 1
		.Col = C_COL2	: .CellType = 1	: .TypeHAlign = 2
		.Row = 1
		.Col = C_COL3	: .CellType = 1	: .TypeHAlign = 2
		.Row = 1
		.Col = C_COL4	: .CellType = 1		: .TypeHAlign = 2 
		.Row = 1 
		.Col = C_SUM	: .CellType = 1	: .TypeHAlign = 2   
		.Row = .Maxrows

		ret = .AddCellSpan(C_H1, C_Row1-1, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row1, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row2, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row3, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row4, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row5, 1, 8)
		ret = .AddCellSpan(C_H1+1, C_Row5, 1, 2)
		ret = .AddCellSpan(C_H1+1, C_Row7, 1, 2)
		ret = .AddCellSpan(C_H1+1, C_Row9, 1, 2)
		ret = .AddCellSpan(C_H1, C_Row11, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row12, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row13, 2, 1)
		ret = .AddCellSpan(C_H1, C_Row14, 2, 1)
		
		.Col =C_COL1	:	.Row =C_Row5	:	.TypeVAlign = 2 :.Col =C_COL2:	.TypeVAlign = 2:.Col =C_COL3:	.TypeVAlign = 2:.Col =C_COL4:	.TypeVAlign = 2:.Col =C_sum:	.TypeVAlign = 2
		.Col =C_COL1	:	.Row =C_Row6	:	.TypeVAlign = 2 :.Col =C_COL2:	.TypeVAlign = 2:.Col =C_COL3:	.TypeVAlign = 2:.Col =C_COL4:	.TypeVAlign = 2:.Col =C_sum:	.TypeVAlign = 2
		.Col =C_COL1	:	.Row =C_Row7	:	.TypeVAlign = 2 :.Col =C_COL2:	.TypeVAlign = 2:.Col =C_COL3:	.TypeVAlign = 2:.Col =C_COL4:	.TypeVAlign = 2:.Col =C_sum:	.TypeVAlign = 2

	End With 
	
	
	
End Sub

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �������� ���� : �޽�����������.
	
	
	if wgConfirmFlg = "Y" then    'Ȯ���� 
	   Exit function
	end if   
	dim TCol
	 wgRefDoc = GetDocRef(sCoCd,sFiscYear, sRepType, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf

    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
    '����� ���α׷� 
    Dim arrW1 ,arrW2 ,  arrW3, arrW4, arrW5, arrW6, iRow, iCol
	call CommonQueryRs("COL1,COL2, COL3, COL4, COL5, COL6","dbo.ufn_TB_23B_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	'If lgF0 = "" Then	 Exit Function
	
	arrW1 = Split(lgF0, chr(11))
	arrW2 = Split(lgF1, chr(11))
	arrW3 = Split(lgF2, chr(11))
	arrW4 = Split(lgF3, chr(11))
	arrW5 = Split(lgF4, chr(11))
	arrW6 = Split(lgF5, chr(11))

	With frm1.vspdData
		.Redraw = False
		
		lgBlnFlgChgValue = True
		ggoSpread.Source = frm1.vspdData
			   
		      For iCol = 3 To  UBound(arrW1) + 2
	                 .Row = 1 '�������� 
					
				 TCol=iCol-3
			      .Col = iCol : .text  = arrW1(TCol)
			       Call vspdData_Change(iCol , 1)
			    
			  
	              .Row = 2 '�����ݾ� 
			      .Col = iCol : .text  = unicdbl(arrW2(TCol))
			       Call vspdData_Change(iCol , 2 )
			    
			
	              .Row = 4 ' ������ش�ݾ� 
			     ' .Col = iCol : .text  = unicdbl(arrW3(TCol))
				  .Col = iCol : .text  = unicdbl(arrW2(TCol))
			       Call vspdData_Change(iCol , 4 )
			    
				
	              .Row = 9 '5�����ʰ��� (����)
			      .Col = iCol : .text  = unicdbl(arrW4(TCol))
			      Call vspdData_Change(iCol , 9 )
			      
	              .Row = 10 '5�����ʰ��� (�и�)
			      .Col = iCol : .text  = unicdbl(arrW5(TCol))
			      Call vspdData_Change(iCol , 10 )
			
	          
			      
			 Next   
			 
		'16ȣ ������ (3)  ��꼭�� ���Աݾ��� �հ�ݾ� �Է�   
		call CommonQueryRs("w6","dbo.ufn_TB_23B_2_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    	 arrW1 = replace(lgF0, chr(11),"")
    	 frm1.txtw6.value =  unicdbl(arrW1)
		
			
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
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()
	Call InitVariables                                                      <%'Initializes local global variables%>

    Call FncQuery
    
End Sub

'============================================  ��������� �Լ�  ====================================
Sub FncSumCal( )
    lgBlnFlgChgValue = True
	' frm1.txtW6.value= unicdbl(Frm1.txtw2.value) + unicdbl(Frm1.txtw4.value)
	 frm1.txtW2.value= unicdbl(Frm1.txtw6.value) - unicdbl(Frm1.txtw4.value)

End Sub

Function  Verification()

	Dim IntRetCD
	dim strWhere

	'zzzzzzz
       'if  unicdbl(frm1.txtW5.value) > unicdbl(frm1.txtW6.value) then
             ' IntRetCD = DisplayMsgBox("WC0010", parent.VB_INFORMATION, "��Ÿ���Աݾ�", "�Ѽ��Աݾ�") 
             ' Exit Function
          'end if
         
  

	Verification = True	
End Function


'==========================================================================================

'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


Sub txtw2_Change( )
    lgBlnFlgChgValue = True
   'Call FncSumCal

End Sub


Sub txtw4_Change( )
    lgBlnFlgChgValue = True
   Call FncSumCal

End Sub

Sub txtw6_Change( )
    lgBlnFlgChgValue = True
   Call FncSumCal

End Sub



Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    if Row <> 1 then
		Dim dblSum, dblW5, dblW7
		dim dblW9, dblW11
		Dim dblW3, dblW13
	
		With frm1.vspdData
	
			Select Case Col
				Case C_Col1, C_Col2, C_Col3, C_Col4 
					if Row =5 or  Row = 7 or  Row =9 or  Row = 11 then 
					   .Col = Col	: .Row = c_row5	: dblW5= UNiCDbl(.Value)
					   .Col = Col	: .Row = c_row7	: dblW7= UNiCDbl(.Value)
					   .Col = Col	: .Row = c_row9	: dblW9= UNiCDbl(.Value)
					   .Col = Col	: .Row = c_row11: dblW11= UNiCDbl(.Value)
					   
					   	.Row = C_Row12
						.Col = Col	: .value = UNiCDbl(dblW5 + dblW7+ dblW9 +dblW11 )
					    Call vspdData_Change(Col , C_Row12 )
					
					end if
					if Row = C_ROW3 OR Row = C_ROW13 then 
					   .Col = Col	: .Row = C_ROW3	: dblW3= UNiCDbl(.Value)
					   .Col = Col	: .Row = C_ROW13	: dblW13= UNiCDbl(.Value)

					   	.Row = C_Row14
						.Col = Col	: .value = UNiCDbl(dblW3 + dblW13 )
					    Call vspdData_Change(Col , C_Row14 )
					
					end if
					
				    Call FncSumSheet(frm1.vspdData, Row, C_COL1, C_Col4, true, Row, C_SUM, "W")
					
			End Select
	
		End With
	end if	
		lgBlnFlgChgValue = True
  '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    If Frm1.vspdData.CellType =  Parent.SS_CELL_TYPE_FLOAT Then
        If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
           Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
        End If
    End If
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
          ' ggoSpread.SSSort Col               'Sort in ascending
          ' lgSortKey = 2
       Else
          ' ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
          ' lgSortKey = 1
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
	
     Call MakeKeyStream("X")
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
	
    
	
	 if Verification = False then Exit Function
	  Call MakeKeyStream("X")

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

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '��: Processing is NG
    
    
    <%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    
    
    If lgIntFlgMode <>  parent.OPMD_UMODE  Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '��: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '��: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If
	Call MakeKeyStream("X")

    If DbDelete= False Then
       Exit Function
    End If												                  '��: Delete db data

    FncDelete=  True                                                              '��: Processing is OK
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
    If lgBlnFlgChgValue = true Then
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
            strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
			strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
			strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey                 '��: Next key tag
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status

	DbDelete = False			                                                 '��: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	strVal = BIZ_PGM_ID & "?txtMode=" &  parent.UID_M0003                                '��: Delete
	strVal = strVal     & "&txtKeyStream="       & lgKeyStream  
	strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey

	Call RunMyBizASP(MyBizASP, strVal)                                           '��: Run Biz logic
	DbDelete = True                                                              '��: Processing is NG

End Function



Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '-----------------------
    'Reset variables area
    '-----------------------
	Call InitData2 ' �׸��� 

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
		Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	End If
	lgBlnFlgChgValue = False
	'frm1.vspdData.focus			
End Function





'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
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
		                                          strVal = strVal & "D"  &  Parent.gColSep
		       Case Else
		                                          strVal = strVal & "U"  &  Parent.gColSep
	       End Select
		 
			For lCol = 2 To lMaxCols
				.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
			Next		
				
				
				strVal = strVal & lRow &  Parent.gColSep &  Parent.gRowSep

		Next
		
        frm1.txtSpread.value         =  strDel & strVal
		frm1.txtMode.value           =  Parent.UID_M0002
		frm1.txtFlgMode.value        =  lgIntFlgMode
		frm1.txtKeyStream.value      =  lgKeyStream
	    'frm1.txtMaxRows.value       =  lGrpCnt - 1

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
					<TD WIDTH=870 valign=top >
					   
					      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>1.���Աݾ׸� </LEGEND>
	
									<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
									   
										
										
										<TR>
												<TD CLASS="TD51" align =center >
													(1)���� 
												</TD>
											
												<TD CLASS="TD51" align =center  >
													(2)�Ϲݼ��Աݾ� 
												</TD>
												<TD CLASS="TD51" align =center  >
													(3)Ư�������ڰ��ŷ�  
												</TD>
												<TD CLASS="TD51" align =center >
													(4)�հ�((2)+(3))
												</TD>
											
										</TR>
									
										<TR>   
										       <TD CLASS="TD51" align =center width = 5%>
													(5)�ݾ� 
												</TD>
											
												
												<TD CLASS="TD61" align =center width = 15% >
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW2" name=txtW2 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="(2)�Ϲݼ��Աݾ�" tag="24X2" width = 100% ></OBJECT>');</SCRIPT>
												</TD>
											

												<TD CLASS="TD61" align =center colspan=1 width = 15% >
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW4" name=txtW4 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="21X2Z" width = 100% ></OBJECT>');</SCRIPT>
												</TD>
											
										
										
												<TD CLASS="TD61" align =center width = 15% >
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtW6" name=txtW6 CLASS=FPDS140 title=FPDOUBLESINGLE ALT="" tag="24X2Z" width = 100% ></OBJECT>');</SCRIPT>
												</TD>

											
										</TR>
									
	
						
									</TABLE>
									
						   </FIELDSET>		


						   
						   			
					</TD>
				</TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">


<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

