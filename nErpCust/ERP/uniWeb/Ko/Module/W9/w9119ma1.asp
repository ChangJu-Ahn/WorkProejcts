
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================

'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ������8ȣ�����ŷ����� 
'*  3. Program ID           : w9119ma1
'*  4. Program Name         : w9119ma1.asp
'*  5. Program Desc         : ������8ȣ�����ŷ����� 
'*  6. Modified date(First) : 2005/03/18
'*  7. Modified date(Last)  : 2006/02/02
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
<SCRIPT LANGUAGE="VBScript"  SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "w9119ma1"
Const BIZ_PGM_ID = "w9119mb1.asp"											 '��: �����Ͻ� ���� ASP�� 

Dim C_W1
Dim C_W1_NM
Dim C_W1_NM2
Dim C_W2
Dim C_W3
Dim C_W4
Dim C_W5
Dim C_W6
Dim C_W7
Dim C_W8
Dim C_W9
Dim C_W10



Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()
    C_W1				= 1
    C_W1_NM				= 2
    C_W1_NM2			= 3
    C_W2				= 4
    C_W3				= 5
    C_W4				= 6
    C_W5				= 7
    C_W6				= 8
    C_W7				= 9
    C_W8				= 10
    C_W9				= 11
    C_W10				= 12
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
    ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
    
	.ReDraw = false


    
    .MaxCols = C_W10 + 1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols														'��: ����� �� Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData

     ggoSpread.SSSetEdit		C_W1,		"�ڵ�", 10,,,10,1
	 ggoSpread.SSSetEdit		C_W1_NM,	"", 5,,,50,1
	 ggoSpread.SSSetEdit		C_W1_NM2,	"", 20,,,50,1
	 ggoSpread.SSSetFloat	    C_W2,		"1"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W3,		"2"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W4,		"3"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W5,		"4"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W6,		"5"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W7,		"6"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W8,		"7"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W9,		"8"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 ggoSpread.SSSetFloat	    C_W10,		"9"	, 20, Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  
	 
	 
	


	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_W1,C_W1,True)
	Call ggoSpread.SSSetColHidden(C_W3,C_W10,True)
    ggoSpread.SSSetSplit2(3)
	' �׸��� ��� ��ħ ���� 


	
					

	
	.ReDraw = true

	Call SetSpreadLock 

    End With   
End Sub


'============================================  �׸��� �Լ�  ====================================



Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_W1, -1, C_W1_NM2

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

	iMaxRows = 18 ' �ϵ��ڵ��Ǵ� ��� 
	With frm1.vspdData
		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData
		
		ggoSpread.InsertRow , iMaxRows

		iRow = 0
		
      for irow = 1 to iMaxRows
		 .Row = iRow
		.Col = C_W1		: .value = iRow
	  NEXT	

		.Redraw = True
		


		
		
		
		
		Call InitData2
		
		Call SetSpreadLock
	End With	

End Sub

 ' -- DBQueryOk ������ �ҷ��ش�.
Sub InitData2()
	Dim iRow  , iCol

	With frm1.vspdData
		.Redraw = False

		iRow = 0
		iRow = iRow + 1 : .Row = iRow    :  .TypeVAlign = 2 :  .rowheight(iRow) = 18
		
		.Col = C_W1_NM	: .value = "  (1)���θ�(��ȣ)"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
	  
		iRow = iRow + 1 : .Row = iRow   :.TypeVAlign = 2 :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = "  (2)�����ּ�"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		
		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2  :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = "  (3)����(������)"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		
		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2  :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = "       (�����ڵ�)"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		

		iRow = iRow + 1 : .Row = iRow    :.TypeVAlign = 2  :  .rowheight(iRow) = 18
		.Col = C_W1_NM	: .value = " (4)�����ΰ��� ����"
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		
		
		

		iRow = iRow + 1 : .Row = iRow     :.TypeVAlign = 2
		.Col = C_W1_NM	: .value = " (5)�հ�((6)+(7))"
		 ggoSpread.SpreadLock	C_W1_NM2  ,iRow , .Maxrows -1 ,iRow
		Call .AddCellSpan(C_W1_NM, iRow ,2, 1) 
		
		
        
		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		 Call .AddCellSpan(C_W1_NM, iRow,1, 6) 
		.Col = C_W1_NM	: .text = "��" & vbCr & "��"& vbCrLf & "��"& vbCrLf & "��" : .TypeEditMultiLine = true   :.TypeVAlign = 2
		
		
		.Col = C_W1_NM2	: .value = " (6)�Ұ�"
		ggoSpread.SpreadLock	C_W1_NM2  ,iRow , .Maxrows -1 ,iRow
		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
	
		
		
		.Col = C_W1_NM2	: .value = "�����ڻ� "
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "�����ڻ�"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "�뿪�ŷ�"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "�������"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "��Ÿ"
		
		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		Call .AddCellSpan(C_W1_NM, iRow ,1, 6) 
		 ggoSpread.SpreadLock	C_W1_NM2  ,iRow , .Maxrows -1 ,iRow
		.Col = C_W1_NM	: .text = "��" & vbCr & "��"& vbCrLf & "��"& vbCrLf & "��" : .TypeEditMultiLine = true   :.TypeVAlign = 2
		
		.Col = C_W1_NM2	: .value = " (7)�Ұ�"
		

		iRow = iRow + 1 : .Row = iRow      :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "�����ڻ� "
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "�����ڻ�"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "�뿪�ŷ�"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "�������"
		
		iRow = iRow + 1 : .Row = iRow       :.TypeVAlign = 2
		.Col = C_W1_NM2	: .value = "��Ÿ"
		
		
	    
		
	  for iRow = 1 to  .maxrows 	
			for iCol = C_W2 to  .maxcols - 1
			         Select Case iRow
			               
			      			  Case 1 ,2, 3 ,4 , 5
									Select Case iCol
									          Case C_W2, C_W3, C_W4, C_W5, C_W6  ,C_W7,  C_W8 , C_W9 ,C_W10
									           .row =iRow : .Col = iCol  :.CellType = 1 
									           ggoSpread.SSSetProtected	iCol  ,iRow,iRow
								                
								             
										End Select	
											
						
							
															
					  End Select 		
				Next	
				
		 
		   					  
		Next    
		

	
		
		
		
	End With
End Sub

'============================== ���۷��� �Լ�  ========================================
Function GetRef()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD , iCol ,strSelect,strFrom,strWhere,ii,jj,arrVal1,arrVal2,iRow
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg , i

	
   if   CommonQueryRs2by2("*", "TB_KJ1" , "Co_Cd  = '" & sCoCd & "'and Fisc_year = '" & sFiscYear & "' and Rep_type = '" & sRepType & "'" , lgF2By2)  =false then
	  	 Call DisplayMsgBox("X", "X", "���� 1ȣ�� ���� �Է��� �ּ���", "X")
	    Exit Function  
	end if    
			  
	  
    wgRefDoc = GetDocRef(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
	
	sMesg = wgRefDoc & vbCrLf & vbCrLf
    
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '��: "Will you destory previous data"
	frm1.vspdData.SetSelection C_W2, 1, C_W2, 1
	
	 If IntRetCD = vbNo Then
	 	Exit Function
	 End If
    
    CALL GetRefB
	
	
	
  
	
End Function


Function GetRefB()	' �ݾװ������� ��ũ Ŭ���� 
	Dim sFiscYear, sRepType, sCoCd, IntRetCD , iCol ,strSelect,strFrom,strWhere,ii,jj,arrVal1,arrVal2,iRow
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	Dim sMesg , i

	' �·ε�� ���۷����޽��� �����´�.
	
	 
	if   CommonQueryRs2by2("*", "TB_KJ1" , "Co_Cd  = '" & sCoCd & "'and Fisc_year = '" & sFiscYear & "' and Rep_type = '" & sRepType & "'" , lgF2By2)  =false then
	  	 Call DisplayMsgBox("X", "X", "���� 1ȣ�� ���� �Է��� �ּ���", "X")
	    Exit Function  
	end if 
	 
	 
	 frm1.vspdData.AllowMultiBlocks = True  
	 frm1.vspdData.AddSelection C_W2, 1, C_W2, 5

	

    
    strSelect = "w1,w2,w3,w4,w5,w6,w7,w8,w9,w10"
    strFrom = "dbo.ufn_TB_KJ8_GetRef('" & sCoCd & "', '" & sFiscYear & "', '" & sRepType & "')"
    strWhere = ""
	with frm1.vspdData	
	    .Redraw = False

			If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			Else 

				
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
				
					For iRow =1 to .MaxRows
							
								For ii = 0 to Ubound(arrVal1,1) - 1
										arrVal2 = Split(arrVal1(ii), chr(11))
							
									.Row = iRow :.Col = C_W1 
								  	if   .Value  =  Trim(arrVal2(1)) then  
								  	    .Col = C_W2 : .Value = Trim(arrVal2(2))
								  	    .Col = C_W3 : .Value = Trim(arrVal2(3))
								  	    .Col = C_W4 : .Value = Trim(arrVal2(4))
								  	    .Col = C_W5 : .Value = Trim(arrVal2(5))
								  	    .Col = C_W6 : .Value = Trim(arrVal2(6))
								  	    .Col = C_W7 : .Value = Trim(arrVal2(7))
								  	    .Col = C_W8 : .Value = Trim(arrVal2(8))
								  	    .Col = C_W9 : .Value = Trim(arrVal2(9))
								  	    .Col = C_W10 : .Value = Trim(arrVal2(10))
								  	  
								  	 End if   

							Next	
					Next		
						
			End If
					
				
			
			
			  for iCol = C_W2  to .maxCols - 1 
		            .col = iCol
		            .Row = 1
				
					   
						if  Trim(.text) <> "" then
						     Call ggoSpread.SSSetColHidden(iCol ,iCol+2,False)
					    
						Else
						      Call ggoSpread.SSSetColHidden(iCol ,iCol+2,True)     
						End if 
						
					     
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
    
    Call SetToolbar("1110100000101111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
    
	Call InitData()
	Call GetRefB()
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

End Sub

Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i

	Dim intCnt,IntRetCD


    If Not chkField(Document, "1") Then					'��: This function check indispensable field
       Exit Function
    End If
   Call SetPrintCond(varCo_Cd, varFISC_YEAR, varREP_TYPE)


     StrUrl = StrUrl & "varCo_Cd|"			& varCo_Cd
	 StrUrl = StrUrl & "|varFISC_YEAR|"		& varFISC_YEAR
	 StrUrl = StrUrl & "|varREP_TYPE|"       & varREP_TYPE
	
	
	
     EBR_RPT_ID	    = "W9119OA1"
     ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
     if  strPrintType = "VIEW" then
	     Call FncEBRPreview(ObjName, StrUrl)
     else
	     Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	

    
    call CommonQueryRs("W4,W5,W6"," TB_KJ1 "," CO_CD= '" & varCo_Cd & "' AND FISC_YEAR='" & varFISC_YEAR & "' AND REP_TYPE='" & varREP_TYPE & "' and w4 >'' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if Trim(replace(lgF0,chr(11),"")) <> "" or   Trim(replace(lgF1,chr(11),"")) <> "" or   Trim(replace(lgF2,chr(11),"")) <> ""  then
      	 EBR_RPT_ID	    = "W9119OA11"
      	 ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
           if  strPrintType = "VIEW" then
			   Call FncEBRPreview(ObjName, StrUrl)
		   else
			   Call FncEBRPrint(EBAction,ObjName,StrUrl)
		   end if	
 
    end if
    
    
    
    call CommonQueryRs("W7,W8,W9"," TB_KJ1 "," CO_CD= '" & varCo_Cd & "' AND FISC_YEAR='" & varFISC_YEAR & "' AND REP_TYPE='" & varREP_TYPE & "' and  w7 >'' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if Trim(replace(lgF0,chr(11),"")) <> "" or   Trim(replace(lgF1,chr(11),"")) <> "" or   Trim(replace(lgF2,chr(11),"")) <> ""  then
      	 EBR_RPT_ID	    = "W9119OA12"
      	 ObjName = AskEBDocumentName(EBR_RPT_ID, "ebr")
           if  strPrintType = "VIEW" then
			   Call FncEBRPreview(ObjName, StrUrl)
		   else
			   Call FncEBRPrint(EBAction,ObjName,StrUrl)
		   end if	
 
    end if
    
   

   
     
   

End Function 

'===============

'===========================================================================






Sub vspdData_Change(ByVal Col , ByVal Row )
 dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6   ,IntRetCD , dblSum1 , dblSum2
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
    
	lgBlnFlgChgValue = True
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
    End If
  

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row


	
	With frm1.vspdData
	
		Select Case ROW
			Case 8,9,10,11,12 ,14,15,16,17,18
			     dblSum1 = FncSumSheet(frm1.vspdData , Col, 8, 12 , true, 7 , Col, "V")	' �հ� 
	
			     dblSum2 = FncSumSheet(frm1.vspdData , Col, 14, 18 , true, 13 , Col, "V")	' �հ� 
			     
			     .Col = Col 
			     .Row = 6
			     .value = unicdbl(dblSum1) + unicdbl(dblSum2)
			     Call vspdData_Change(Col , 6)
			     Call vspdData_Change(Col , 7)
			     Call vspdData_Change(Col , 13)
			     
		
		End Select
		


	
	End With
End Sub



Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	Dim strTemp
	Dim intPos1

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        
    
    End With
    
    

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

    Call SetToolbar("1110100000001111")

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
   
    
    Call InitData()

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
        FncDelete = False                                                             '��: Processing is NG
    
    
        
        
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    
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

    If DbDelete= False Then
       Exit Function
    End If												                  '��: Delete db data

    FncDelete=  True                          
    
   lgBlnFlgChgValue = True
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
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim iCol
    
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
			Call SetToolbar("1111100000011111")										<%'��ư ���� ���� %>

		Else
	
		End If
	Else
		Call SetToolbar("1111100000011111")										<%'��ư ���� ���� %>
	End If

          
     With frm1.vspdData
     
        for iCol = C_W2  to .maxCols - 1 


            .col = iCol
            .Row = 1
            
			
			if  Trim(.text) <> "" then
			     Call ggoSpread.SSSetColHidden(iCol ,iCol,False)
			    
			Else
			      Call ggoSpread.SSSetColHidden(iCol ,iCol,True)     
			End if 
				
			     
        Next
        
     End With   
 
    Call InitData2()
	lgBlnFlgChgValue = False
	
    
	'Call SetSpreadLock(TYPE_1)
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
		
		   For lRow = 6 To .MaxRows
    
               .Row = lRow
               .Col = 0
            
				Select Case  .Text
				    Case  ggoSpread.InsertFlag                                      '��: Insert
		

						 strVal = strVal & "C"  &  Parent.gColSep '0
						
							.Col = C_W1     : strVal = strVal & Trim(.Text) &  Parent.gColSep '1
							.Col = C_W2     : strVal = strVal & Trim(.Text) &  Parent.gColSep '2
							.Col = C_W3     : strVal = strVal & Trim(.Text) &  Parent.gColSep '3
							.Col = C_W4     : strVal = strVal & Trim(.Text) &  Parent.gColSep '4
							.Col = C_W5     : strVal = strVal & Trim(.Text) &  Parent.gColSep '5
							.Col = C_W6     : strVal = strVal & Trim(.Text) &  Parent.gColSep '6
							.Col = C_W7     : strVal = strVal & Trim(.Text) &  Parent.gColSep '7
							.Col = C_W8     : strVal = strVal & Trim(.Text) &  Parent.gColSep '8
							.Col = C_W9     : strVal = strVal & Trim(.Text) &  Parent.gColSep '9
							.Col = C_W10     : strVal = strVal & Trim(.Text) &  Parent.gRowSep '10			                         
							
				
				            lGrpCnt = lGrpCnt + 1
				        
				    Case  ggoSpread.UpdateFlag                                      '��: Update
				           strVal = strVal & "U"  &  Parent.gColSep
				       
							.Col = C_W1     : strVal = strVal & Trim(.Text) &  Parent.gColSep '1
							.Col = C_W2     : strVal = strVal & Trim(.Text) &  Parent.gColSep '2
							.Col = C_W3     : strVal = strVal & Trim(.Text) &  Parent.gColSep '3
							.Col = C_W4     : strVal = strVal & Trim(.Text) &  Parent.gColSep '4
							.Col = C_W5     : strVal = strVal & Trim(.Text) &  Parent.gColSep '5
							.Col = C_W6     : strVal = strVal & Trim(.Text) &  Parent.gColSep '6
							.Col = C_W7     : strVal = strVal & Trim(.Text) &  Parent.gColSep '7
							.Col = C_W8     : strVal = strVal & Trim(.Text) &  Parent.gColSep '8
							.Col = C_W9     : strVal = strVal & Trim(.Text) &  Parent.gColSep '9
							.Col = C_W10     : strVal = strVal & Trim(.Text) &  Parent.gRowSep '10			                         
							
					                 
				        lGrpCnt = lGrpCnt + 1
				        
				         Case  ggoSpread.DeleteFlag                                      '��: Update
				         strVal = strVal & "D"  &  Parent.gColSep
				        
							.Col = C_W1        : strVal = strVal & Trim(.Text) &  Parent.gRowSep '1
					                    
				        lGrpCnt = lGrpCnt + 1 
				        
				 
				End Select
				  
							  
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
	strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'��: ��ȸ ���� ����Ÿ 


	Call RunMyBizASP(MyBizASP, strVal)                                           '��: Run Biz logic
	DbDelete = True                                                              '��: Processing is NG

End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
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
									<TD CLASS="TD6"><script language =javascript src='./js/w9119ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
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
									<script language =javascript src='./js/w9119ma1_vaSpread1_vspdData.js'></script>
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
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
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

