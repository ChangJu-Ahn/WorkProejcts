<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b26ma1.asp
'*  4. Program Name         : Query Characteristic
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/07
'*  8. Modified date(Last)  :  
'*  9. Modifier (First)     : Lee Woo Guen33333333
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Grid 1 - Operation
Const BIZ_PGM_QRY_ID1	= "b3b26mb1.asp"								'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY_ID2	= "b3b26mb2.asp"								'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Operation 
Dim C_CharCd			'= 1
Dim C_CharNm			'= 2
Dim C_CharValueDigit	'= 3

' Grid 2(vspdData2) - Operation
Dim C_CharValueCd		'= 1
Dim C_CharValueNm		'= 2
																
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntGrpCount							'Group View Size�� ������ ���� 
Dim lgIntFlgMode							'Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgSortKey1
Dim lgSortKey2

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
         
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvSpdNo)
    
    Call InitSpreadPosVariables(pvSpdNo)
    
    If pvSpdNo = "A" Or pvSpdNo = "*" Then 
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1
			 ggoSpread.Source = frm1.vspdData1
			 ggoSpread.Spreadinit "V20021107", , Parent.gAllowDragDropSpread
			.ReDraw = false
			 
			.MaxCols = C_CharValueDigit + 1 											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit		C_CharCd, 			"����׸�"		, 23
			ggoSpread.SSSetEdit 	C_CharNm,			"����׸��"	, 30
			ggoSpread.SSSetEdit 	C_CharValueDigit,	"��簪�ڸ���"	, 18,1

			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("A")
			
			.ReDraw = true
    
		End With
    End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021107", , Parent.gAllowDragDropSpread
			
			.ReDraw = false
			.MaxCols = C_CharValueNm + 1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit 	C_CharValueCd,	"��簪"		, 20    
			ggoSpread.SSSetEdit 	C_CharValueNm, 	"��簪��"		, 30
						
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
 					
			'ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true 
    
		End With
    End If
    
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadLock(ByVal pvSpdNo)

    With frm1
		Select Case pvSpdNo
			Case "A"
				'--------------------------------
				'Grid 1
				'--------------------------------
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.SpreadLockWithOddEvenRowColor()
		
			Case "B"
				'--------------------------------
				'Grid 2
				'--------------------------------
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadLockWithOddEvenRowColor()
		End Select
    End With
    
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitComboBox()

End Sub


'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Operation 
		C_CharCd			= 1
		C_CharNm			= 2
		C_CharValueDigit	= 3
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Operation
		C_CharValueCd		= 1
		C_CharValueNm		= 2
	End If
		
End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData1
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_CharCd			= iCurColumnPos(1)
			C_CharNm			= iCurColumnPos(2)
			C_CharValueDigit	= iCurColumnPos(3)
			
		Case "B"
			ggoSpread.Source = frm1.vspdData2 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
 			C_CharValueCd		= iCurColumnPos(1)
			C_CharValueNm		= iCurColumnPos(2)
 						
 	End Select
 
End Sub

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenCharCd()  -------------------------------------------------
'	Name : OpenCharCd()
'	Description : Characteristic PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCharCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtCharCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtCharCd.value)	' Characteristic Code
	arrParam(1) = ""							' Characteristic Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field��(0) : "Char_CD"
    arrField(1) = 2 							' Field��(1) : "Char_NM"
    
	iCalledAspName = AskPRAspName("B3B30PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B30PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=610px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		

	IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetCharCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCharCd.focus
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetCharCd()  --------------------------------------------------
'	Name : SetCharCd()
'	Description : Characteristic Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetCharCd(byval arrRet)
	frm1.txtCharCd.Value	= arrRet(0)	
	frm1.txtCharNm.Value	= arrRet(1)
	
	frm1.txtCharCd.focus
	Set gActiveElement = document.activeElement
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet("*")                                               '��: Setup the Spread sheet
  '  Call InitVariables                                                     '��: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    'Call InitComboBox
    Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
    
	frm1.txtCharCd.focus 
	Set gActiveElement = document.activeElement 
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    '----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1
    
 	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		
 		lgOldRow = Row
		
		lgStrPrevKey2 = ""
			
		frm1.vspdData2.MaxRows = 0
			  		
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then
			frm1.vspdData1.Col = 1
			frm1.vspdData1.Row = row
	
			lgOldRow = Row
			
			lgStrPrevKey2 = ""
				
			frm1.vspdData2.MaxRows = 0
			  		
			'frm1.KeyCharCd.value = frm1.vspdData1.Text
			  		
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
			
		End If
	 	'------ Developer Coding part (End)	
 	End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================

Sub vspdData2_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2 
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey2 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey2		'Sort in Descending
 			lgSortKey2 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
 	
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )

End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'==========================================================================================
'   Event Name : vspdData_DragDropBlock
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData2_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )

End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
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
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.ID)
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop)  Then
		If lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False                                                   '��: Processing is NG

    Err.Clear                                                          '��: Protect system from crashing

	If frm1.txtCharCd.value = "" Then
		frm1.txtCharNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")								'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables
    																	'��: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If																'��: Query db data

    FncQuery = True														'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                                   <%'��: Protect system from crashing%>
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                                    <%'��: Protect system from crashing%>
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing
        
    With frm1
    
	    If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID1 & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtCharCd=" & Trim(.hCharCd.value)				'��: ��ȸ ���� ����Ÿ 
			
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode		
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	    Else
			strVal = BIZ_PGM_QRY_ID1 & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtCharCd=" & Trim(.txtCharCd.value)			'��: ��ȸ ���� ����Ÿ 
			
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	    End If

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

	Call SetToolBar("11000000000111")										'��: ��ư ���� ���� 
	
	lgOldRow = 1
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		frm1.vspddata1.Row = 1
		frm1.vspddata1.Col = 1
		frm1.KeyCharCd.value = frm1.vspddata1.Text
							
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement		
		
		lgIntFlgMode = parent.OPMD_UMODE
		
	End If
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 
    Dim strVal
    Dim strCharCd    
    
	frm1.vspdData1.Col = C_CharCd
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow

	strCharCd = Trim(frm1.vspdData1.Text)

    DbDtlQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
        
    With frm1
    
		Call LayerShowHide(1)
    
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & parent.UID_M0001			'��: 
			strVal = strVal & "&txtCharCd=" & Trim(strCharCd)				'��: ��ȸ ���� ����Ÿ 
			
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows			
		Else
					
			strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & parent.UID_M0001			'��: 
			strVal = strVal & "&txtCharCd=" & Trim(strCharCd)				'��: ��ȸ ���� ����Ÿ 
			
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbDtlQuery = True

End Function


Function DbDtlQueryOk()														'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	lgAfterQryFlg = True	
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����׸���ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
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
									<TD CLASS=TD5 NOWRAP>����׸�</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCharCd" SIZE=25 MAXLENGTH=18 tag="11XXXU"  ALT="����׸�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtCharNm" SIZE=40 tag="14"></TD>
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
							<TR HEIGHT="100%">
								<TD WIDTH="50%" colspan=4>
								<script language =javascript src='./js/b3b26ma1_A_vspdData1.js'></script></TD>
								<TD WIDTH="50%" colspan=4>
								<script language =javascript src='./js/b3b26ma1_B_vspdData2.js'></script></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hCharCd" tag="24"><INPUT TYPE=HIDDEN NAME="KeyCharCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hCharValueCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>

