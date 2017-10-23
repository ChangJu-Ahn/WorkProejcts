<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ư��2ȣ���׸�����û�� 
'*  3. Program ID           : W6105MA1
'*  4. Program Name         : W6105MA1.asp
'*  5. Program Desc         : ��Ư��2ȣ���׸�����û�� 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/02/07
'*  8. Modifier (First)     : ȫ���� 
'*  9. Modifier (Last)      : HJo 
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

Const BIZ_MNU_ID		= "W6105MA1"
Const BIZ_PGM_ID		= "W6105MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID	    = "W6105OA1"



' -- �׸��� �÷� ���� 
Dim	C_SEQ_NO

Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W4_VAL
Dim C_W5
Dim C_W6


Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgCurrGrid,IsRunEvents


'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO		= 1
	C_W1		= 2
	C_W2		= 3
	C_W3		= 4
	C_W4		= 5
	C_W4_VAL		= 6
	C_W5		= 7
	C_W6		= 8
	
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


    IsRunEvents = False
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

End Sub

Sub InitSpreadSheet()
	Dim ret, iRow
	

    Call initSpreadPosVariables()  

	'Call AppendNumberPlace("6","3","2")	' -- ����(����)
	
	' 1�� �׸��� 

	With  frm1.vspdData
				
		ggoSpread.Source = frm1.vspdData	
		'patch version
		ggoSpread.Spreadinit "V20021127",,parent.gForbidDragDropSpread     
    
		.ReDraw = false

		.MaxCols = C_W6 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData
        Call AppendNumberPlace("6","3","3")
		ggoSpread.SSSetEdit     C_SEQ_NO, "", 3,,,10,1
		ggoSpread.SSSetEdit     C_W1, "(1)����", 40,,,60,1
		ggoSpread.SSSetEdit     C_W2, "(2)�ٰŹ�����",  20,,,50,1
		ggoSpread.SSSetEdit     C_W3, "(3)�ڵ�", 6,2,,50,1
		ggoSpread.SSSetEdit     C_W4, "(4)������", 8,2,,50,1
		ggoSpread.SSSetFloat    C_W4_Val,"(4) ��������",   15,	    "6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
		ggoSpread.SSSetFloat    C_W5,   "(5)��󼼾�", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,,"0"
		ggoSpread.SSSetFloat    C_W6,"(6)��������",   15,      Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,,,,"0"
					

			
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W4_Val,C_W4_Val,True)
		
		'Call InitSpreadComboBox
		
		.ReDraw = true	
				
	End With 

 
	
	
End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitData()
	
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
     
 
    
    
    
End Sub

Sub SpreadInitData()
    ' �׸��� �ʱ� ����Ÿ���� 
    Dim arrW1, arrW2, arrW3,arrW4, iMaxRows, iRow, iMinorCnt, sMinorCd, ret , strFrom,strW2,strW1 ,iSpanRow

		strFrom = "  ufn_TB_Configuration('W1069' ,'" & C_REVISION_YM & "')  a "
		
		call CommonQueryRs(" a.minor_cd ,a.minor_nm,a.reference_1 ,a.reference_2",strFrom," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        arrW1		= Split(lgF0, Chr(11))
		arrW2		= Split(lgF1, Chr(11))
		arrW3		= Split(lgF2, Chr(11))
		arrW4		= Split(lgF3, Chr(11))

    
 		iMaxRows = UBound(arrW1)
	
		With frm1.vspdData
			.Redraw = False
			
			ggoSpread.Source = frm1.vspdData
			
			ggoSpread.InsertRow , iMaxRows
            .Redraw = True
		
			' �迭�� �׸��忡 ���� 
			
				Call SetSpreadLock()
			For iRow = 1 To iMaxRows
				
				.Row = iRow
				.Col = C_SEQ_NO	: .text = arrW1(iRow-1)
				.Col = C_W1		: .text = arrW2(iRow-1)
				.Col = C_W2		: .text = arrW3(iRow-1)
				.Col = C_W3 	: .text = arrW4(iRow-1)
				.Col = C_W1	
				
				
		         if  Trim(.text) = "" then
			   	     ggoSpread.SpreadUnLock	C_W1, iRow, C_W2, iRow
			   	     ggoSpread.SpreadUnLock	C_W4, iRow, C_W4, iRow
			   	     ggoSpread.SSSetMask    C_W4,"������(4)", 10, ,"99%",iRow
			   	 end if    				
			 Next							
		end With
End Sub

Sub SetSpreadLock()
dim i
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	

				ggoSpread.SpreadLock C_SEQ_NO, -1, C_W4, -1	' ��ü ���� 
				'ggoSpread.SpreadLock C_W6, -1, C_W6, -1	' ��ü ���� 
			  	ggoSpread.SpreadLock C_SEQ_NO, .MaxRows, C_W6,  .MaxRows	' ��ü ���� 
				
			

	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData

			
	End With	
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)

       
End Sub


Sub SetSpreadTotalLine()
	Dim iRow, i


		ggoSpread.Source = frm1.vspdData
		With frm1.vspdData
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1 : .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				'ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With

End Sub 

' �ش� �׸��忡�� ����Ÿ�������� 
Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With frm1.vspdData
		.Col = pCol	: .Row = pRow : GetGrid = UNICDbl(.Value)
	End With
End Function

' �ش� �׸��忡�� ����Ÿ�������� 
Function PutGrid(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	With frm1.vspdData
		.Col = pCol	: .Row = pRow : .Value = pVal
	End With
End Function

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �׸���1�� �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2,arrW3, arrW4,arrW5, iMaxRows, sTmp,jj,arrW6
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

	' ����� ��ġ�� �˷��� 
	Dim iCol, iRow
	

	With frm1.vspdData


	   .Redraw = False	
	   .AddSelection C_W4, 1, C_W4, .maxrows' -- �������� ������ �߰��Ҷ� 
	

	
		IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
		Call ggoOper.LockField(Document, "N") 
		.SetSelection iCol, 1, iCol, 1
		
		If IntRetCD = vbNo Then
			 Exit Function
		End If
	.Redraw = True
	End With



	IntRetCD = CommonQueryRs("SEQ_NO,W1 ,W4 ,W4_VAL,W5,W6 "," dbo.ufn_TB_JT2_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		arrW3		= Split(lgF2, chr(11))
		arrW4		= Split(lgF3, chr(11))
		arrW5		= Split(lgF4, chr(11))
		arrW6		= Split(lgF5, chr(11))
		iMaxRows	= UBound(arrW1)

		With frm1.vspdData
		
				For iRow = 1 To .Maxrows -1

						For   jj = 0 to iMaxRows
		
						    .Row = iRow :.Col = C_SEQ_NO 
						    if    trim(.text)  =  Trim(arrW1(jj)) then  
						          .Row = iRow :.Col = C_W1 
						          if  trim(.Value) = "" then
								      .Col = C_W1 : .text = arrW2(jj)
								   end if   
						          .Col = C_W4 : .value = arrW3(jj)
						          .Col = C_W4_VAL : .text = arrW4(jj)
						          .Col = C_W5 : .text = arrW5(jj)
						          .Col = C_W6 : .text = arrW6(jj)
						           'Call vspdData_Change(C_W4_VAL,iRow)
						           
						    end  if
						NEXt
				Next
				
				  Call FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows-1, true, .MaxRows, C_W5, "V")	' ���� ���� �հ� 
                  Call FncSumSheet(frm1.vspdData, C_W6, 1, .MaxRows-1, true, .MaxRows, C_W6, "V")	' ���� ���� �հ� 
            
		
		End With
		
		'Call SetReCalc1
	End If
	
	
	frm1.vspdData.focus
	lgBlnFlgChgValue = True
	
	
	
	
End Function

Sub ReCalcGrid()

	Dim dblW
	
End Sub



'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         

    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>

    Call InitVariables                                                      <%'Initializes local global variables%>
   
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
 	
	Call InitComboBox	
	Call InitData
	Call SpreadInitData

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

Sub cboREP_TYPE_onChange()	' �Ű������ �ٲٸ�..
	'Call GetFISC_DATE
End Sub



'============================================  �׸��� �̺�Ʈ   ====================================



'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange( ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change( ByVal Col , ByVal Row )
	Dim dblSum, dblSum141,IROW,IntRetCD,str07Row,dblAmt , dblW5,dblRate,dblW4
	Dim sFiscYear, sRepType, sCoCd
	lgBlnFlgChgValue= True ' ���濩�� 
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col

    If frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(frm1.vspdData.text) < UNICDbl(frm1.vspdData.TypeFloatMin) Then
        frm1.vspdData.text = frm1.vspdData.TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	' --- �߰��� �κ� 
	With frm1.vspdData


		Select Case Col
			Case C_W4
				 .Row = Row
				 .Col = Col
			     dblW4 = unicdbl(.value)
			    .Col = C_W4_Val : .value = dblW4 / 100
			     Call vspdData_Change(C_W4_Val,Row)
			Case C_W4_Val

			     Call vspdData_Change(C_W6,Row)
            Case C_W5
                 
			     .Col = C_W4_Val : .Row = Row : dblRate = .text
			     .Col = C_W5	 : .Row = Row : dblW5 = .text 
			    ' .Col = C_W6	 : .Row = Row :.Value   = unicdbl(dblRate) * unicdbl(dblW5)
                 Call FncSumSheet(frm1.vspdData, C_W5, 1, .MaxRows-1, true, .MaxRows, C_W5, "V")	' ���� ���� �հ� 
                 Call FncSumSheet(frm1.vspdData, C_W6, 1, .MaxRows-1, true, .MaxRows, C_W6, "V")	' ���� ���� �հ� 
                 ggoSpread.UpdateRow .MaxRows
            Case C_W6
                 Call FncSumSheet(frm1.vspdData, C_W6, 1, .MaxRows-1, true, .MaxRows, C_W6, "V")	' ���� ���� �հ� 
		          ggoSpread.UpdateRow .MaxRows
		End Select
	

	
	End With
	
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)

    'Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
     ' If Row <= 0 Then
   '     ggoSpread.Source = frm1.vspdData
   '     If lgSortKey = 1 Then
   '         ggoSpread.SSSort Col
   '         lgSortKey = 2
   '     Else
   '         ggoSpread.SSSort Col,lgSortKey
   '         lgSortKey = 1
   ''     End If    
   '     Exit Sub

	frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange( ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick( ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub



Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

	ggoSpread.Source = frm1.vspdData
End Sub    

Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange( ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
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

Sub vspdData_ButtonClicked( ByVal Col, ByVal Row, Byval ButtonDown)

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
    If lgBlnFlgChgValue Then
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
    Call InitVariables		
    Call SpreadInitData											<%'Initializes local global variables%>
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
	Call SpreadInitData
	
    Call SetToolbar("1100100000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If frm1.vspdData.ActiveRow > 0 Then
			frm1.vspdData.focus
			frm1.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor lgCurrGrid, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

			frm1.vspdData.Col = C_W13
			frm1.vspdData.Text = ""
    
			frm1.vspdData.Col = C_W3
			frm1.vspdData.Text = ""
			
			frm1.vspdData.Col = C_W4
			frm1.vspdData.Text = ""
			
			frm1.vspdData.Col = C_W5
			frm1.vspdData.Text = ""
			
			frm1.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 

End Function


Function FncInsertRow(ByVal pvRowCnt) 

End Function

Function FncDeleteRow() 

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
Dim IntRetCD, iRow
	
	FncExit = False
    If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
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
        'strVal = strVal     & "&txtMaxRows="         & frm1.vspdData.MaxRows            
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function
		
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
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
	       	Call SetToolbar("110000000000111")			
		End If
	Else
		Call SetToolbar("1101100000000111")								<%'��ư ���� ���� %>
	End If
	lgBlnFlgChgValue = False

	'Call SetSpreadLock(TYPE_1)

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
    Dim strVal, strDel, sTmp
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1

    
		With frm1.vspdData
	
			ggoSpread.Source = frm1.vspdData
			lMaxRows = .MaxRows : lMaxCols = .MaxCols
			
			' ----- 1��° �׸��� 
			For lRow = 1 To .MaxRows

    
				.Row = lRow	: sTmp = "" : .Col = 0
		    
				  ' ��� �׸��� ����Ÿ ����     
				  If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
						For lCol = 1 To lMaxCols
							Select Case lCol
								'Case C_W31
								'	.Col = lCol : strVal = strVal & Trim(.Text) &  Parent.gColSep
								Case Else
									.Col = lCol : sTmp = sTmp & Trim(.Value) &  Parent.gColSep
							End Select
						Next
						sTmp = sTmp & Trim(.Text) &  Parent.gRowSep
				  End If  

				.Col = 0
				Select Case .Text
					Case  ggoSpread.InsertFlag                                      '��: Insert
				                                       strVal = strVal & "C"  &  Parent.gColSep & sTmp
				    Case  ggoSpread.UpdateFlag                                      '��: Update
				                                       strVal = strVal & "U"  &  Parent.gColSep & sTmp
				    Case  ggoSpread.DeleteFlag                                      '��: Update
				                                       strDel = strDel & "D"  &  Parent.gColSep & sTmp
				End Select

			Next
							   
		End With


		
	Frm1.txtSpread.value      = strDel & strVal

    Frm1.txtFlgMode.value     = lgIntFlgMode
	Frm1.txtMode.value        =  Parent.UID_M0002
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Dim iRow
	
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
					<TD WIDTH=* align=right><a href="vbscript:GetRef">�ݾ� �ҷ�����</A>  </TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w6105ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
                        <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="10">&nbsp;[��û����]				
								</TD>
							</TR>
							<TR>
								<TD >
									<script language =javascript src='./js/w6105ma1_vspdData_vspdData0.js'></script>
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
			<TABLE CLASS="TB3" CELLSPACING=0>
			    
		
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" style="display:'none'"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtw2_val" tag="24">
<INPUT TYPE=HIDDEN NAME="txtW5_A_val" tag="24">
<INPUT TYPE=HIDDEN NAME="txtW5_B_val" tag="24">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

