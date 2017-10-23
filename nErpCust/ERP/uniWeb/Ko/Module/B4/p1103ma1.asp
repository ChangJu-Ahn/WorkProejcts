<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1103ma1.asp
'*  4. Program Name         : Į����Ÿ�� ��� 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Mr  KimGyoungDon
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p1103mb1.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "p1103mb2.asp"											'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const C_SHEETMAXROWS = 30

Dim C_ClnrType
Dim C_ClnrTypeNm

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop   
Dim IsQuery       

<!-- #Include file="../../inc/lgvariables.inc" -->	


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

	
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
	lgSortKey = 1
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================= 2.2.3 InitSpreadSheet() ======================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()

	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData		
		ggoSpread.Spreadinit "V20021121", , Parent.gAllowDragDropSpread

		.ReDraw = False
	
		.MaxCols = C_ClnrTypeNm + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit 	C_ClnrType, "Į���� Ÿ��",		20,,,2,2
		ggoSpread.SSSetEdit 	C_ClnrTypeNm, "Į���� Ÿ�Ը�",	96,,,40

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True

		Call SetSpreadLock 

    End With
    
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()
	
	C_ClnrType = 1
	C_ClnrTypeNm = 2
	
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

			C_ClnrType    = iCurColumnPos(1)
			C_ClnrTypeNm  = iCurColumnPos(2)
			
	End Select

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
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

'	Parent.gMouseClickStatus = "SPC"   
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("1101011111")
   
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col					'Sort in Ascending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortKey = 1
		End If
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
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	
	
'   If Button = 2 And Parent.gMouseClickStatus = "SPC" Then
   If Button = 2 And gMouseClickStatus = "SPC" Then

      gMouseClickStatus = "SPCR"
   End If
   
   
   
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



'================= 2.2.4 SetSpreadLock() ================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    With frm1

    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_ClnrType, -1, C_ClnrType
	ggoSpread.SSSetRequired 	C_ClnrTypeNm,	-1
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1
	.vspdData.ReDraw = True
	
	End With
End Sub

'=========================== 2.2.5 SetSpreadColor() =====================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1
    
    .vspdData.ReDraw = False

	ggoSpread.SSSetRequired 	C_ClnrType,		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_ClnrTypeNm,	pvStartRow, pvEndRow
	
    .vspdData.ReDraw = True
    
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
'Sub InitComboBox()

'End Sub

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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCalType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Į���� Ÿ�� �˾�"			' �˾� ��Ī 
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtClnrType.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "Į���� Ÿ��"					' TextBox ��Ī 
	
    arrField(0) = "CAL_TYPE"						' Field��(0)
    arrField(1) = "CAL_TYPE_NM"						' Field��(1)
    
    arrHeader(0) = "Į���� Ÿ��"				' Header��(0)
    arrHeader(1) = "Į���� Ÿ�Ը�"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCalType(arrRet)
	End If	
    
    Call SetFocusToDocument("M")
	frm1.txtClnrType.focus
    
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetCalType()  -----------------------------------------------
'	Name : SetCalType()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)

	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1)
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
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
'    Call InitComboBox
    Call InitVariables
    Call SetToolbar("11001100000011")										'��: ��ư ���� ���� 
    
    frm1.txtClnrType.focus
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData 

		If Row >= NewRow Then
		    Exit Sub
		End If

	'----------  Coding part  -------------------------------------------------------------

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, 0) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
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

    FncQuery = False															'��: Processing is NG

    Err.Clear																    '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then                   '��: Check If data is chaged

		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtClnrType.value = "" Then
		frm1.txtClnrTypeNm.value = ""
	End If
		
	Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	Call ggoSpread.ClearSpreadData
    Call InitVariables
  
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     					'��: Query db data

    FncQuery = True																'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 

     On Error Resume Next														'��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 

     On Error Resume Next														'��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 

    Dim IntRetCD 
    
    FncSave = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged

        IntRetCD = DisplayMsgBox("900001","X","X","X")            '��: Display Message(There is no changed data.)

        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then					  '��: Check required field(Multi area)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     														'��: Save db data
    
    FncSave = True																'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement 
    ggoSpread.Source = frm1.vspdData

	frm1.vspdData.EditMode = True	
	frm1.vspdData.ReDraw = False
    ggoSpread.CopyRow
    
    frm1.vspdData.Col = C_ClnrType
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Text = ""
    
    frm1.vspdData.ReDraw = True
    
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function	
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo															'��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 


	Dim IntRetCD
	Dim imRow
		
	On Error Resume Next
	
	FncInsertRow = False
   
	If IsQuery= True then
		If frm1.vspdData.maxrows >= 0 Then 
			Call SetToolbar("11001111001011")										'��: ��ư ���� ���� 
		End if				 
	Else
		If frm1.vspdData.maxrows >= 0 Then 
			Call SetToolbar("11001101001011")										'��: ��ư ���� ���� 
		End if				 
	End If
'	imRow = parent.AskSpdSheetAddRowCount()

'	IF imRow = "" Then
'		Exit Function
'	End If
	
	If IsNumeric(Trim(pvRowCnt)) Then
		
		imRow = CInt(pvRowCnt)
		
	Else
		
		imRow= AskSpdSheetAddRowCount()
		
		If imRow = "" Then
			Exit Function
		End If
		
	End If

	
	With frm1

		.vspdData.ReDraw = False	
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow .vspdData.ActiveRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
		.vspdData.ReDraw = True
		.vspdData.EditMode = True
   
    End With
    	
	'----------------- Developer Coding Part (Start) --------------------
	'----------------- Developer Coding Part (End)   --------------------
	
    Set gActiveElement = document.activeElement 
		
	IF Err.number = 0 Then
		FncInsertRow = True
	End IF
   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt, i
    
    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
	If frm1.vspdData.maxrows < 1 Then Exit Function
    
    With frm1.vspdData 
    
    .focus
    Set gActiveElement = document.activeElement
     
    ggoSpread.Source = frm1.vspdData 
	lDelRows = ggoSpread.DeleteRow
    
    End With
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

    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 

    Call parent.FncFind(parent.C_MULTI, False)                                        '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

	Dim IntRetCD
	FncExit = False
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 

    If ggoSpread.SSCheckChange = True Then                   '��: Check If data is chaged
'		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")		
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")		

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

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

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    DbQuery = False
    IsQuery = False
    LayerShowHide(1)
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtClnrType=" & Trim(.hClnrType.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtClnrType=" & Trim(.txtClnrType.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    IsQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	    
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call SetToolbar("11001111001111")
	
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
on error resume next
err.Clear
    Dim lRow        
    Dim lGrpCnt     
   	Dim strVal
	Dim strDel
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt
	Dim iColSep
	
    DbSave = False                                                          '��: Processing is NG
    
    LayerShowHide(1)
		
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    iValCnt = 0
    iDelCnt = 0
    
    ReDim TmpBufferVal(0)
    ReDim TmpBufferDel(0)
    
    lGrpCnt = 1
    
	iColSep = Parent.gColSep
	
    '-----------------------
    'Data manipulate area
    '-----------------------

    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag												'��: �ű� 
				
				strVal = ""
				
				strVal = strVal & "C" & iColSep 									'��: C=Create
				
                .vspdData.Col = C_ClnrType	'1
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_ClnrTypeNm	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = strVal
                
                iValCnt = iValCnt + 1
                
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag
				
				strVal = ""
				
				strVal = strVal & "U" & iColSep										'��: U=Update
				
                .vspdData.Col = C_ClnrType	'1
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_ClnrTypeNm	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = strVal
                
                iValCnt = iValCnt + 1
                                
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.DeleteFlag												'��: ���� 
			
				strDel = ""
				
				strDel = strDel & "D" & iColSep	
				
                .vspdData.Col = C_ClnrType	'1
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

                ReDim Preserve TmpBufferDel(iDelCnt)
				
                TmpBufferDel(iDelCnt) = strDel

                iDelCnt = iDelCnt + 1

                lGrpCnt = lGrpCnt + 1

        End Select
                
    Next
	
	iTotalStrVal = Join(TmpBufferVal, "")
	iTotalStrDel = Join(TmpBufferDel, "")
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = iTotalStrDel & iTotalStrVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True																	'��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()																	'��: ���� ������ ���� ���� 

	Call InitVariables
    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--#########################################################################################################
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Į����Ÿ�Ե��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>Į���� Ÿ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=5 MAXLENGTH=2 tag="11XXXU" ALT="Į���� Ÿ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=40 tag="14" ALT="Į���� Ÿ�Ը�"></TD>
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
							<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/p1103ma1_I179984492_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hClnrType" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
