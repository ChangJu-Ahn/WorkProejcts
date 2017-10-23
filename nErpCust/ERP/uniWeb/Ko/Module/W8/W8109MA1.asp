<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��13ȣ�����Ư����������󰨸鼼���հ�ǥ 
'*  3. Program ID           : W3101MA1
'*  4. Program Name         : W3101MA1.asp
'*  5. Program Desc         : ��13ȣ�����Ư����������󰨸鼼���հ�ǥ 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/02/07
'*  8. Modifier (First)     : ȫ���� 
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

Const BIZ_MNU_ID		= "W8109MA1"
Const BIZ_PGM_ID		= "W8109MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "W8101OA1"

Const TYPE_1	= 0		' �׸��带 �������� ���� ��� 


' -- �׸��� �÷� ���� 
Dim	C_W1_CD
Dim C_W2_CD
Dim C_W1
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W7


Dim IsOpenPop  
Dim gSelframeFlg        
Dim lgStrPrevKey2
Dim lgCurrGrid, lgvspdData(2), IsRunEvents


'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
	C_W1_CD		= 1
	C_W1		= 2
	C_W2_CD		= 3
	C_W2		= 4
	C_W3		= 5
	C_W4		= 6
	C_W7		= 7
	
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

    lgCurrGrid = TYPE_1
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
	
	Set lgvspdData(TYPE_1)		= frm1.vspdData0


		
    Call initSpreadPosVariables()  

	'Call AppendNumberPlace("6","3","2")	' -- ����(����)
	
	' 1�� �׸��� 

	With lgvspdData(TYPE_1)
				
		ggoSpread.Source = lgvspdData(TYPE_1)	
		'patch version
		ggoSpread.Spreadinit "V20041222_" & TYPE_1,,parent.gForbidDragDropSpread    
    
		.ReDraw = false

		.MaxCols = C_W7 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols									'��: ����� �� Hidden Column
		.ColHidden = True    
		
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		ggoSpread.SSSetEdit		C_W1_CD,	"����", 10,,,15,1
		ggoSpread.SSSetEdit		C_W2_CD,	"�ڵ�", 4,,,15,1
		ggoSpread.SSSetEdit		C_W1,		"(1)����", 10,,,50,1
		ggoSpread.SSSetEdit		C_W2,		"(2)���鼼��", 35,,,50,1
		ggoSpread.SSSetEdit		C_W3,		"(3)�ٰ�", 21,,,50,1
		
		ggoSpread.SSSetFloat	C_W4,		"(4)���鼼��"	, 15, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec	 ,,,,"0"
		ggoSpread.SSSetEdit		C_W7,		"���"	, 30,,,50,1

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_W1_CD,C_W1_CD,True)
		
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
     
    Call SetDefaultVal
    
End Sub
  
Sub SetDefaultVal()     
    call CommonQueryRs(" reference_1,reference_2"," ufn_TB_Configuration('w2015','" & C_REVISION_YM & "') "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)					' ���չ��ο� ���� ���μ��� 
    frm1.txtw2.value = replace(lgF1,chr(11),"")
    frm1.txtw2_val.value = replace(lgF0,chr(11),"")
    
     call CommonQueryRs(" reference_1,reference_2"," ufn_TB_Configuration('w2001','" & C_REVISION_YM & "') "," minor_cd= '1'  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  ' ���μ��� - 1������ 
		frm1.txtW5_A.value = replace(lgF1,chr(11),"")
		frm1.txtW5_A_val.value = replace(lgF0,chr(11),"")
    
     call CommonQueryRs(" reference_1,reference_2"," ufn_TB_Configuration('w2001','" & C_REVISION_YM & "') "," minor_cd= '2'  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  ' ���μ��� - 1���ʰ� 
        frm1.txtW5_B.value = replace(lgF1,chr(11),"")
		frm1.txtW5_B_val.value = replace(lgF0,chr(11),"")
End Sub
    


Sub SpreadInitData()
    ' �׸��� �ʱ� ����Ÿ���� 
    Dim arrW1_CD, arrW1, arrW2, arrW2_1, iMaxRows, iRow, iMinorCnt, ret , strFrom,arrW1_span , arrW_type,strW2,strW1 ,iSpanRow


		strFrom = "  ufn_TB_Configuration('w1074','" & C_REVISION_YM & "') a "																			    '��Ư������󰨸鼼�ױ���(H)
        strFrom = strFrom &" left join  ufn_TB_Configuration('w1075','" & C_REVISION_YM & "') b on  b.reference_3 = a.minor_cd"                             '��Ư������󰨸鼼�׳���(D)

		call CommonQueryRs(" a.minor_cd ,b.minor_cd , a.minor_nm,b.minor_nm, b.reference_1 ,b.reference_2",strFrom," 1=1 ORDER BY a.minor_cd, b.reference_4 ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        arrW_type   = Split(lgF0, Chr(11))
		arrW1_CD	= Split(lgF1, Chr(11))
		arrW1_span	= Split(lgF2, Chr(11))
		arrW1		= Split(lgF3, Chr(11))
		arrW2		= Split(lgF4, Chr(11))
		arrW2_1		= Split(lgF5, Chr(11))
    
 		iMaxRows = UBound(arrW1_CD)
	
		With lgvspdData(TYPE_1)
			.Redraw = False
			
			ggoSpread.Source = lgvspdData(TYPE_1)
			
			ggoSpread.InsertRow , iMaxRows
            .Redraw = True
		
			' �迭�� �׸��忡 ���� 			

					For iRow = 1 To iMaxRows
		
						.Row = iRow
						.Col = C_W1_CD	: .value = arrW_type(iRow-1)
						.Col = C_W2_CD	: .value = arrW1_CD(iRow-1)
						.Col = C_W1: .value = arrW1_span(iRow-1)
						.Col = C_W2 	: .value = arrW1(iRow-1)
						.Col = C_W3		: .value = arrW2(iRow-1)
						.Col = C_W7  	: .value = arrW2_1(iRow-1)
						
					     .Col = C_W1_CD	:.Row = iRow 
						 
						 if strW1 =  Trim(.Text) and iRow <> 1 then                          '������ ���� �� ������ �������� ��ģ�� 
		
						    ret = .AddCellSpan(C_W1	, iSpanRow  , 1, Irow - iSpanRow +1)	
						   
						 else
						    .Col = C_W1_CD	:.Row = iRow : strW1  = Trim(.Text)  
						     iSpanRow = iRow
						 end if

						  	
			
						.Col = C_W2_CD	:.Row = iRow : strW2  = Trim(.Text)
						 if left(strW2,2) = "00"  or strW2 = "" then                           '�ڵ��׸��� ���� ���� ���� ��ģ�� 
						    ret = .AddCellSpan(C_W1	, iRow, 3, 1)	
						    ggoSpread.SpreadLock C_w1, .Row, C_W7,  .Row	' ��ü ���� 
						 end if
					 Next						
		end With				
		
		Call SetSpreadLock(TYPE_1)
End Sub

Sub SetSpreadLock(pType)
dim i
	With lgvspdData(pType)
	
		ggoSpread.Source = lgvspdData(pType)	

		Select Case pType
			Case TYPE_1 
				ggoSpread.SpreadLock C_W1, -1, C_W3, -1	
				ggoSpread.SpreadLock C_W7, -1, C_W7, -1
				ggoSpread.SpreadLock C_w1, .MaxRows, C_W7,  .MaxRows	
				
				for i = 1 to .maxrows  -1
				    .Col = C_W2 : .Row = i
	
					if trim(.Text) = "" then
					    	ggoSpread.SpreadunLock C_W2,i , C_W4, i	'
					    	
					end if
				next 
				
			
				
		End Select
		
	End With	
End Sub

' InsertRow/Copy �Ҷ� ȣ��� 
Sub SetSpreadColor(Byval pType, ByVal pvStartRow, ByVal pvEndRow)

	With lgvspdData(pType)
		ggoSpread.Source = lgvspdData(pType)	

			
	End With	
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case pvSpdNo
       Case TYPE_1
            ggoSpread.Source = frm1.vspdData0
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_W1_CD		= iCurColumnPos(1)
            C_W2_CD		= iCurColumnPos(2)
            C_W3		= iCurColumnPos(3)
            C_W4		= iCurColumnPos(4)
            C_W7		= iCurColumnPos(5)
            
 
       
    End Select    
End Sub


Sub SetSpreadTotalLine()
	Dim iRow, i

	For i = TYPE_1 To TYPE_1
		ggoSpread.Source = lgvspdData(i)
		With lgvspdData(i)
			If .MaxRows > 0 Then
				.Row = .MaxRows
				.Col = C_W1 : .CellType = 1	: .Text = "��"	: .TypeHAlign = 2
				'ggoSpread.SSSetProtected -1, .MaxRows, .MaxRows
			End If
		End With
	Next
End Sub 

' �ش� �׸��忡�� ����Ÿ�������� 
Function GetGrid(Byval pType, Byval pCol, Byval pRow)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : GetGrid = UNICDbl(.Value)
	End With
End Function

' �ش� �׸��忡�� ����Ÿ�������� 
Function PutGrid(Byval pType, Byval pCol, Byval pRow, Byval pVal)
	With lgvspdData(pType)
		.Col = pCol	: .Row = pRow : .Value = pVal
	End With
End Function

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �׸���1�� �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD, arrW1, arrW2, iMaxRows, sTmp,   jj
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg

	' �·ε�� ���۷����޽��� �����´�.
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)

	sMesg = wgRefDoc & vbCrLf & vbCrLf

	' ����� ��ġ�� �˷��� 
	Dim iCol, iRow
	

	With frm1.vspdData0
		iCol = .ActiveCol	: iRow = .ActiveRow

	 .Redraw = False	
		FOR  IROW = 1 TO  .MAXROWS -1 
		   
		        .row = IROW 
		        .col = C_W2_CD 
                .AllowMultiBlocks = True  
               
		  	Select Case Trim(.text)
			
				   
				case "121" ,"122", "123", "126", "127", "128", "129", "125", "131", "132", "133", "134", "135","136","137","138", "140", "141", "142", "139"
					 
			         .AddSelection C_W4, IROW, C_W4, IROW' -- �������� ������ �߰��Ҷ� 
			         ggoSpread.UpdateRow IROW
			End Select
		NEXT 
		

	
		IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
		Call ggoOper.LockField(Document, "N") 
		.SetSelection iCol, 1, iCol, 1
		
		If IntRetCD = vbNo Then
			 Exit Function
		End If
	.Redraw = True
	End With



	IntRetCD = CommonQueryRs("W1,W2 "," dbo.ufn_TB_13_GetRef_200603('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If IntRetCD = True Then
		arrW1		= Split(lgF0, chr(11))
		arrW2		= Split(lgF1, chr(11))
		iMaxRows	= UBound(arrW1)

		With frm1.vspdData0
		
				For iRow = 1 To .Maxrows -1

						For   jj = 0 to iMaxRows
		
						    .Row = iRow :.Col = C_W2_CD
						    if    trim(.Value)  =  Trim(arrW1(jj)) then  
						          .Row = iRow
						          .Col = C_W4       : .value = arrW2(jj)

						    end  if
						NEXt
				Next
	
		
		End With
		
		Call SetReCalc1
	End If
	
	
	frm1.vspdData0.focus
	lgBlnFlgChgValue = True
	
	
	
	
End Function

Sub SetReCalc1()

	Call FncSumSheet(frm1.vspdData0, C_W4, 8, frm1.vspdData0.MaxRows-1, true, frm1.vspdData0.MaxRows, C_W4, "V")
	ggoSpread.UpdateRow frm1.vspdData0.MaxRows
	
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

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
 	
	Call InitComboBox	
	Call InitData
	Call SpreadInitData
    'Call SetDefaultVal
    
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


Sub txtw1_Change()	
    frm1.txtw3.value = unicdbl(frm1.txtw1.value) *  unicdbl(trim(frm1.txtw2_val.value))  
    frm1.txtw3_SUM.value = frm1.txtw3.value 

    lgBlnFlgChgValue = true
End Sub

Sub txtw3_Change()	
    frm1.txtw7.value = unicdbl(frm1.txtw6_SUM.value) -  unicdbl(frm1.txtw3_SUM.value)

End Sub

Sub txtw4_A_Change()	
    frm1.txtw6_A.value = unicdbl(frm1.txtw4_A.value) *  unicdbl(trim(frm1.txtW5_A_val.value))  
    frm1.txtw4_SUM.value = unicdbl(frm1.txtw4_A.value) +  unicdbl(trim(frm1.txtW4_B.value))
     lgBlnFlgChgValue = true
End Sub


Sub txtw4_B_Change()	
    frm1.txtw6_B.value = unicdbl(frm1.txtw4_B.value) *  unicdbl(trim(frm1.txtW5_B_val.value)) 
     frm1.txtw4_SUM.value = unicdbl(frm1.txtw4_A.value) +  unicdbl(trim(frm1.txtW4_B.value)) 
     lgBlnFlgChgValue = true 
End Sub


Sub txtw6_a_Change()	
     frm1.txtw6_SUM.value = unicdbl(frm1.txtw6_A.value) +  unicdbl(trim(frm1.txtW6_B.value))  

End Sub
Sub txtw6_b_Change()	
    Call txtw6_a_Change

End Sub

Sub txtw6_SUM_Change()	
    Call txtw3_Change

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

Sub vspdData0_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	lgCurrGrid = TYPE_1
	vspdData_ButtonClicked lgCurrGrid, Col, Row, ButtonDown
End Sub



'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
Sub vspdData_ComboSelChange(Index, ByVal Col, ByVal Row)

End Sub

Sub vspdData_Change(Index, ByVal Col , ByVal Row )
	Dim dblSum, dblSum141,IROW,IntRetCD,str07Row,dblAmt , dbl120Amt
	Dim sFiscYear, sRepType, sCoCd , dblMonth
	lgBlnFlgChgValue= True ' ���濩�� 
    lgvspdData(lgCurrGrid).Row = Row
    lgvspdData(lgCurrGrid).Col = Col

    If lgvspdData(Index).CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(lgvspdData(Index).text) < UNICDbl(lgvspdData(Index).TypeFloatMin) Then
         lgvspdData(Index).text = lgvspdData(Index).TypeFloatMin
      End If
    End If
	
    ggoSpread.Source = lgvspdData(Index)
    ggoSpread.UpdateRow Row
 

	' --- �߰��� �κ� 
	With lgvspdData(Index)

	If Index = TYPE_1 Then	'1�� �׸� 
		Select Case Col
			Case C_W4
			
				sCoCd		= frm1.txtCO_CD.value
				sFiscYear	= frm1.txtFISC_YEAR.text
				sRepType	= frm1.cboREP_TYPE.value

		
		       if Row < 8 then		
				
						        dblSum =  FncSumSheet(lgvspdData(TYPE_1), Col, 1, 7, False, 8, Col, "V")	' ���� ���� �հ� 

								IntRetCD = CommonQueryRs("isnull(W4,0) , isnull(W120,0)"," dbo.ufn_TB_13_1_GetRef('" & sCoCd & "','" & sFiscYear & "','" & sRepType & "')", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

								if IntRetCD  =TRUE  then
									dblAmt = unicdbl(lgF0)           '����ǥ�� 
									dbl120Amt = unicdbl(lgF1)		 '���� ���� 
								Else
								   dblAmt= 0
								   dbl120Amt =0
								END IF	
							   		
							 
							    Call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear  & "' AND REP_TYPE='" & sRepType  & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
							          dblMonth =  DateDiff("m", CDate(lgF0), CDate(lgF1)) + 1
								.col = C_W4 :.row= 8
							
								 '1���̻� =>  {[(����ǥ�� + (1~7 ����)] * 12/������� -1��) * 1���̻��� } * �������/12 + (1��* 1��̸�����) *(�������/12) - ���⼼�� 
						    if  (unicdbl(dblAmt) + unicdbl(dblSum)) * (12 /unicdbl(dblMonth))  > 100000000 then  
						        .text =  (( (unicdbl(dblAmt) + unicdbl(dblSum)) *(12/unicdbl(dblMonth)) -100000000) * unicdbl(frm1.txtW5_B_val.value) ) * (unicdbl(dblMonth)/12) +  (100000000  * unicdbl(frm1.txtW5_A_val.value)) *(unicdbl(dblMonth) /12) - unicdbl(dbl120Amt)
						       
						    else '1������ =>  ����ǥ�� 
						        .text =  (( unicdbl(dblAmt) + unicdbl(dblSum))  * (12/unicdbl(dblMonth)) * unicdbl(frm1.txtW5_A_val.value)) *  unicdbl(dblMonth)/12  - unicdbl(dbl120Amt)
						    end if 
						    
						     ggoSpread.UpdateRow lgvspdData(Index).row
				End if	
                   Call FncSumSheet(lgvspdData(TYPE_1), Col, 8, .MaxRows-1, true, .MaxRows, Col, "V")
                        ggoSpread.UpdateRow lgvspdData(Index).MaxRows
	                 
					
					
		End Select
	

	End If
	
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
   	    
    If Row <= 0 Then
       ggoSpread.Source = lgvspdData(Index)
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
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
    ggoSpread.Source = lgvspdData(Index)
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos(Index)
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

Sub vspdData_ButtonClicked(Index, ByVal Col, ByVal Row, Byval ButtonDown)

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
	For i = TYPE_1 To TYPE_1
		ggoSpread.Source = lgvspdData(i)
		If ggoSpread.SSCheckChange = True Then
			blnChange = True
			Exit For
		End If
    Next
    
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
        'strVal = strVal     & "&txtMaxRows="         & lgvspdData(lgCurrGrid).MaxRows            
		
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
			Call SetToolbar("1101100000001111")										<%'��ư ���� ���� %>

		Else
	
		End If
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
    

    
		With lgvspdData(TYPE_1)
	
			ggoSpread.Source = lgvspdData(TYPE_1)
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
	
	For iRow = TYPE_1 To TYPE_1
	
		lgvspdData(iRow).MaxRows = 0
		ggoSpread.Source = lgvspdData(iRow)
		ggoSpread.ClearSpreadData
	Next
    Call SetDefaultVal 	
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
									<TD CLASS="TD6"><script language =javascript src='./js/w8109ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE  width = 100%  bgcolor=#ffffff BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=*>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD HEIGHT="10">&nbsp;1. ���μ� ���� ���� ����					
										</TD>
									</TR>
									<TR>
										<TD >
											<script language =javascript src='./js/w8109ma1_vspdData0_vspdData0.js'></script>
										</TD>
									</TR>
									
								</TABLE>
								</TD>
							</TR>
							<tr>
						 <TD WIDTH=900 valign=top HEIGHT=* >
					   
								      <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>2.���չ��ε� ���鼼�� </LEGEND>
												<TABLE width = 100% bgcolor = #696969  border = 0 cellpadding = 1 cellspacing = 1 ID="Table1">
												   
													
													<TR>
														<TD CLASS="TD51" align =center rowspan = 2   width=10%>
															(1)���μ� ����ǥ�� 
														</TD>
														
													    <TD CLASS="TD51" align =center rowspan = 2   width=10%>
															(2)����Ư��<br>���ѹ�<br>��72������ 
														</TD>
														<TD CLASS="TD51" align =center rowspan = 2  width=10%>
															(3)���⼼��<br>((1)x(2))
														</TD>
														<TD CLASS="TD51" align =center  colspan = 2 width=15%>
															(4)����ǥ�� 
														</TD>
														<TD CLASS="TD51" align =center  rowspan = 2 width=10%>
															(5)���μ��� ��55<br>������ 
														</TD>
														<TD CLASS="TD51" align =center rowspan = 2  width=10%>
															(6)���⼼�� 
														</TD>
														<TD CLASS="TD51" align =center  rowspan = 2  width=10%>
															(7)���鼼��((6)-(3))
														</TD>
													</TR>
														<TR>
														<TD CLASS="TD51" align =center width = 5% >
															���� 
														</TD>
														
													    <TD CLASS="TD51" align =center  >
															�ݾ� 
														</TD>
													</TR>
													<TR>
													    <TD CLASS="TD61" align =center  rowspan = 2>
															<script language =javascript src='./js/w8109ma1_txtw1_txtw1.js'></script>
														</TD>
														
														 <TD CLASS="TD61" align =center rowspan = 2  width=5%>
															<INPUT type="text" id=txtw2 name=txtw2  TAG="24X"  size =3 maxlength=3>
														</TD>
														 <TD CLASS="TD61" align =center rowspan = 2>
															<script language =javascript src='./js/w8109ma1_txtW3_txtW3.js'></script>
														</TD>
														 <TD CLASS="TD61" align =center>
															1������ 
														</TD>
														<TD CLASS="TD61" align =left >
															<script language =javascript src='./js/w8109ma1_txtW4_A_txtW4_A.js'></script>
														</TD>
														 <TD CLASS="TD61" align =center>
															<INPUT type="text" id=txtw5_A name=txtw5_A    size =3  tag="24X" >
														</TD>
														<TD CLASS="TD61" align =left   >
															<script language =javascript src='./js/w8109ma1_txtW6_A_txtW6_A.js'></script>
														</TD>
														 <TD CLASS="TD61" align =center   rowspan = 2 >
														
														</TD>
													 </tr>
													 <TR>
													
														
														
													    <TD CLASS="TD61" align =center>
															1���ʰ� 
														</TD>
														
														<TD CLASS="TD61" align =left >
															<script language =javascript src='./js/w8109ma1_txtW4_B_txtW4_B.js'></script>
														</TD>
														 <TD CLASS="TD61" align =center>
															<INPUT type="text" id=txtw5_B name=txtw5_B  size =3   tag="24X" >
														</TD>
														<TD CLASS="TD61" align =left >
															<script language =javascript src='./js/w8109ma1_txtW6_B_txtW6_B.js'></script>
														</TD>
													
													 </tr>
													 <TR>
													    <TD CLASS="TD61" align =center  colspan = 2 >
															�հ� 
														</TD>
														
														 <TD CLASS="TD61" align =center>
															<script language =javascript src='./js/w8109ma1_txtW3_Sum_txtW3_Sum.js'></script>
														</TD>
														 <TD CLASS="TD61" align =center >
															�հ� 
														</TD>
														<TD CLASS="TD61" align =left >
															<script language =javascript src='./js/w8109ma1_txtW4_Sum_txtW4_Sum.js'></script>
														</TD>
														 <TD CLASS="TD61" align =center >
															�հ� 
														</TD>
														<TD CLASS="TD61" align =left >
															<script language =javascript src='./js/w8109ma1_txtW6_Sum_txtW6_Sum.js'></script>
														</TD>
														<TD CLASS="TD61" align =left >
															<script language =javascript src='./js/w8109ma1_txtW7_txtW7.js'></script>
														</TD>
													</tr>	
						
												</TABLE>
									   </FIELDSET>				
									   			
								</TD>
							</tr>
                        </TABLE>
                        </DIV>
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

