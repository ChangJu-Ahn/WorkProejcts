<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : List Material Valuation
'*  3. Program ID           : I2311ma1.asp
'*  4. Program Name         : ���庰  ��� ���� 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'
'*  7. Modified date(First) : 2000/05/3
'*  8. Modified date(Last)  : 2005/02/17
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              : Tracking No addition
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/18 : ..........
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'            1. �� �� �� 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
' ���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                

'******************************************  1.2 Global ����/��� ����  ***********************************
' 1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_ID = "i2311mb1.asp"           

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Dim	C_ItemCode           
Dim	C_ItemName     
Dim	C_Spec         
Dim	C_ItemUnit     
Dim	C_Location     
Dim	C_Qty          
Dim	C_Amount       
Dim	C_Price        
Dim	C_PriceInd     
Dim	C_PrevQty      
Dim	C_PrevAmount   
Dim	C_PrevPrice    
Dim	C_PrevPriceInd
Dim C_TrackingNo

Dim lgStrPrevToKey
dim DbChkFlag


 '==========================================  1.2.2 Global ���� ����  =====================================
' 1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
' 2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
lgIntFlgMode = Parent.OPMD_CMODE 
lgIntGrpCount = 0                         
'---- Coding part--------------------------------------------------------------------
lgStrPrevToKey =	1
lgLngCurRows = 0                          

End Sub


'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End if 
 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemAcct.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20050217", , Parent.gAllowDragDropSpread
	
	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_PrevPriceInd+1         
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")		
				  
		ggoSpread.SSSetEdit		C_ItemCode,		"ǰ��",			18
		ggoSpread.SSSetEdit		C_ItemName,		"ǰ���",		25
		ggoSpread.SSSetEdit		C_Spec,			"�԰�",			20
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No",	20
		ggoSpread.SSSetEdit		C_ItemUnit,		"����",			10,2
		ggoSpread.SSSetEdit		C_Location,		"Location",		20
		ggoSpread.SSSetFloat	C_Qty,			"��������",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_Amount,		"�����ݾ�",	15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_Price,		"�ܰ�",			15, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec 
		ggoSpread.SSSetEdit		C_PriceInd,		"�ܰ�����",		10,2
		ggoSpread.SSSetFloat	C_PrevQty,		"����������", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_PrevAmount,	"�������ݾ�", 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat	C_PrevPrice,	"�����ܰ�",		15, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec 
		ggoSpread.SSSetEdit		C_PrevPriceInd,	"�����ܰ�����",	13,2  
		  
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		.ReDraw = true

		Call SetSpreadLock
		ggoSpread.SSSetSplit2(2)
	End With

End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCode     = 1
	C_ItemName     = 2
	C_Spec         = 3
	C_TrackingNo   = 4
	C_ItemUnit     = 5
	C_Location     = 6
	C_Qty          = 7
	C_Amount       = 8
	C_Price        = 9
	C_PriceInd     = 10
	C_PrevQty      = 11
	C_PrevAmount   = 12
	C_PrevPrice    = 13
	C_PrevPriceInd = 14
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCode     = iCurColumnPos(1)
		C_ItemName     = iCurColumnPos(2)
		C_Spec         = iCurColumnPos(3)
		C_TrackingNo   = iCurColumnPos(4)
		C_ItemUnit     = iCurColumnPos(5)
		C_Location     = iCurColumnPos(6)
		C_Qty          = iCurColumnPos(7)
		C_Amount       = iCurColumnPos(8)
		C_Price        = iCurColumnPos(9)
		C_PriceInd     = iCurColumnPos(10)
		C_PrevQty      = iCurColumnPos(11)
		C_PrevAmount   = iCurColumnPos(12)
		C_PrevPrice    = iCurColumnPos(13)
		C_PrevPriceInd = iCurColumnPos(14)
	End Select

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
 With frm1
  .vspdData.ReDraw = False
  ggoSpread.SpreadLockWithOddEvenRowColor()
  .vspdData.ReDraw = True
 End With
End Sub


'------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�" 
	arrParam(1) = "B_PLANT"    
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""   
	arrParam(5) = "����"   
 
	arrField(0) = "PLANT_CD" 
	arrField(1) = "PLANT_NM" 
 
	arrHeader(0) = "����"  
	arrHeader(1) = "�����"  

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)  
		frm1.txtPlantNm.Value    = arrRet(1) 
		frm1.txtPlantCd.focus 
	End If 
 
End Function

 '------------------------------------------  OpenItem()  -------------------------------------------------
' Name : OpenItem()
' Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam(5), arrField(6)
	 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	 
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Trim(FilterVar(frm1.txtPlantCd.Value," ","S")), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X", "X", "X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus 
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
	 
	If Trim(frm1.txtItemAcct.Value) <> "" then
		If  CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD = " & Trim(FilterVar(frm1.txtItemAcct.Value," ","S")), _
						  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		   
			Call DisplayMsgBox("169952",vbOKOnly, "x", "x")   
			frm1.txtItemAcctNm.value = ""
			frm1.txtItemAcct.focus 
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtItemAcctNm.value = lgF0(0)
	End If
	 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	 
	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtItemCd.Value) 
	arrParam(2) = ""
	arrParam(3) = Trim(frm1.txtItemAcct.Value)
	 
	arrField(0) = 1  
	arrField(1) = 2  
	arrField(2) = 9  
	arrField(3) = 6  

	iCalledAspName = AskPRAspName("B1B11PA3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0) 
		frm1.txtItemNm.value = arrRet(1)
		frm1.txtItemCd.focus
	End If 
End Function

'------------------------------------------  OpenItemAcct()  --------------------------------------------------
' Name : OpenItemAcct()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAcct()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function
 
 IsOpenPop = True

 arrParam(0) = "ǰ����� �˾�"    
 arrParam(1) = "B_MINOR"      
 arrParam(2) = Trim(frm1.txtItemAcct.Value) 
 arrParam(3) = ""       
 arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & ""  
 arrParam(5) = "ǰ�����"   
 
 arrField(0) = "MINOR_CD"      
 arrField(1) = "MINOR_NM"      
 
 arrHeader(0) = "ǰ�����"     
 arrHeader(1) = "������"      
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtItemAcct.focus
	Exit Function
 Else
	frm1.txtItemAcct.Value		= arrRet(0)
	frm1.txtItemAcctNm.Value	= arrRet(1)
	frm1.txtItemAcct.focus
 End If 
End Function


'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
' Name : OpenTrackingNo()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Tracking No."	
	arrParam(1) = "S_SO_TRACKING"				
	arrParam(2) = Trim(frm1.txtTrackingNo.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "Tracking No."			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking_No"		
    arrHeader(1) = "ǰ��"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
	End If	
End Function


'------------------------------------------  OpenStockRef()  -------------------------------------------------
' Name : OpenStockRef()
' Description : OnhandStock/OnhandStock detail Reference (����OpenStockRef p4212ra ���� i2312ra ����)
'--------------------------------------------------------------------------------------------------------- 

Function OpenStockRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1
	Dim Param2
	Dim Param3
	 
	If IsOpenPop = True Then Exit Function

	Param1 = Trim(frm1.txtPlantCd.value)
	if Param1 = "" then
		Call DisplayMsgBox("169901","X", "X", "X") 
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	ggoSpread.Source = frm1.vspdData

	If frm1.vspdData.value = "" Then
		Call DisplayMsgBox("900002","X","X","X")
		frm1.txtItemAcct.focus 
		Exit Function
	Else
		With frm1.vspdData 
			If .MaxRows = 0 Then
				Exit Function
			else
			    .Col = C_ItemCode
			    .Row = .ActiveRow
			     Param2 = Trim(.Text )
			     
			     .Col = C_TrackingNo
			    .Row = .ActiveRow
			     Param3 = Trim(.Text )
			End If 
		End With
	
		if Param2 = "" then
			Call DisplayMsgBox("169903","X", "X", "X")
			Exit Function
		End If
	End if
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2312RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2312RA1","x")
		IsOpenPop = False
		Exit Function
	End If
 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")  
	 
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.focus
		Exit Function
	End If 
End Function

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
 
    Call LoadInfTB19029           
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call InitSpreadSheet          
    Call InitVariables            
    
     '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
	Call SetToolbar("11000000000011")         
    
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

     If OldLeft <> NewLeft Then
         Exit Sub
     End If
     
     If CheckRunningBizProcess = True Then
   Exit Sub
  End If
     
 if DbChkFlag = true then exit sub
    
 if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then 

		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End if
	
 End if 
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
   
	Set gActiveSpdSheet = frm1.vspdData
   
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		
			lgSortKey = 1
		End If
	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	Dim iColumnName
   
	If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()

	FncQuery = False                                                      
 
	Err.Clear                                                             
 
	'-----------------------
	'Erase contents area
	'-----------------------
 
	Call InitVariables               
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then      
	 Exit Function
	End If

	'-----------------------
	'�����ڵ尡 �ִ� �� üũ 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Trim(FilterVar(frm1.txtPlantCd.Value," ","S")), _
	 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  
	 Call DisplayMsgBox("125000","X", "X", "X")    
	 frm1.txtPlantNm.value = ""
	 frm1.txtPlantCd.focus 
	 Exit function
	End If
		lgF0 = Split(lgF0,Chr(11))
	   frm1.txtPlantNm.value = lgF0(0)
	   
	'-----------------------
	'ǰ������ڵ尡 �ִ��� üũ 
	'-----------------------
	If  CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD = " & Trim(FilterVar(frm1.txtItemAcct.Value," ","S")), _
	 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  
	 Call DisplayMsgBox("169952",vbOKOnly, "x", "x")   
	 frm1.txtItemAcctNm.value = ""
	 frm1.txtItemAcct.focus 
	 Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	   frm1.txtItemAcctNm.value = lgF0(0)
	   
	frm1.txtItemNm.value = ""
	If frm1.txtItemCd.value <> "" Then
	 If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
	  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
	 
	  lgF0 = Split(lgF0,Chr(11))
	  frm1.txtItemNm.value = lgF0(0)
	 End If
	End If
	    
	'-----------------------
	'Query function call area
	'----------------------- 
	Call SetToolbar("11000000000111") 
 
	If DbQuery = False Then               
	 Exit Function
	End if
 
	FncQuery = True               
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint()
 Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================

Function FncExcel()
Call parent.FncExport(Parent.C_MULTI)        
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , True)        
End Function



'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit() 
    FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery()
 Dim LngLastRow
 Dim LngMaxRow
 Dim StrNextKey
 Dim StrNextKey2
 Dim strFlag
    
	If frm1.RadioOutputType.rdoCase1.Checked Then
		strFlag = "Y"
	Else
		strFlag = "N"
	End if
 
 Call LayerShowHide(1)

 DbQuery = False
 
 Err.Clear                                                               
 
 Dim strVal
 With frm1
  if lgIntFlgMode = Parent.OPMD_UMODE Then
   strVal = BIZ_PGM_ID	& "?txtPlantCd="	& .hPlantCd.value				& _
						"&txtItemCd="		& .hItemCd.value				& _
						"&txtAccntCd="		& .hAccntCd.value				& _
						"&txtTrackingNo="	& .hTrackingNo.value			& _
						"&txtFlag="			& strFlag						& _
						"&lgStrPrevToKey="	& lgStrPrevToKey				& _
						"&txtMaxRows="		& .vspdData.MaxRows
  Else
   strVal = BIZ_PGM_ID	& "?txtPlantCd="    & Trim(.txtPlantCd.value)		& _
						"&txtItemCd="		& Trim(.txtItemCd.value)		& _
						"&txtQryFrItemCd="	& Trim(.txtItemCd.value)		& _
						"&txtAccntCd="		& Trim(.txtItemAcct.value)		& _
						"&txtTrackingNo="	& Trim(.txtTrackingNo.value)	& _
						"&txtFlag="			& strFlag						& _
						"&lgStrPrevToKey="	& lgStrPrevToKey				& _
						"&txtMaxRows="		& .vspdData.MaxRows
  End if
 End With
 Call RunMyBizASP(MyBizASP,strVal)
 
 
 DbQuery = True

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()              
 
'-----------------------
'Reset variables area
'-----------------------
frm1.vspdData.focus
lgIntFlgMode = Parent.OPMD_UMODE 


End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE <%=LR_SPACE_TYPE_00%>>
			<TR>
				<TD <%=HEIGHT_TYPE_00%> >
				</TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE  <%=LR_SPACE_TYPE_10%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
									<TR>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���庰 �����ȸ</font></TD>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
									</TR>
								</TABLE>
							</TD>
							<TD WIDTH=* align=right><A href="vbscript:OpenStockRef()">�����Ȳ</A></TD>     
							<TD WIDTH=10></TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% CLASS="Tab11">
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD <%=HEIGHT_TYPE_02%> >
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD">
									<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%> >
										<TR>
											<TD CLASS="TD5" NOWRAP>����</TD>      
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=6 NAME="txtPlantCd" MAXLENGTH="4"  tag="13XXXU" ALT = "����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=27 MAXLENGTH=40 tag="14"></TD>    
											<TD CLASS="TD5" NOWRAP>ǰ��</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT = "ǰ��" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=40 tag="14"></TD>       
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcct" SIZE=6 MAXLENGTH=2  tag="12XXXU" ALT="ǰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=27 MAXLENGTH=50 tag="14"></TD>
											<TD CLASS="TD5" NOWRAP>Tracking No.</TD>      
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>������������</TD>      
											<TD CLASS="TD6" NOWRAP >
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">��������</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">��ǰ��</LABEL>
												</TD>
											<TD CLASS="TD5" NOWRAP></TD>
											<TD CLASS="TD6" NOWRAP></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=* valign=top>
									<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%> >
										<TR>
											<TD HEIGHT="100%">
											<script language =javascript src='./js/i2311ma1_I361804252_vspdData.js'></script></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_01%> >
					</TD>
				</TR>
				<TR HEIGHT=20 >
					<TD>
						<TABLE <%=LR_SPACE_TYPE_30%> >
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=<%=BizSize%>>
						<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1">
						</IFRAME>
					</TD>
				</TR>
			</TABLE>
			<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
			<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hAccntCd" tag="24" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24" TABINDEX="-1">
		</FORM>
			<DIV ID="MousePT" NAME="MousePT">
				<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
			</DIV>
	</BODY>
</HTML>

