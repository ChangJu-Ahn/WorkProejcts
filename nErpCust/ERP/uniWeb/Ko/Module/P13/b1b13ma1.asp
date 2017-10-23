<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Master Production Scheduling
'*  3. Program ID			: b1b13ma1.asp
'*  4. Program Name			: ��üǰ���� 
'*  5. Program Desc			:
'*  6. Business ASP List	: 
'*  7. Modified date(First)	: 2000/11/03
'*  8. Modified date(Last)	: 2002/12/16
'*  9. Modifier (First)		: 
'* 10. Modifier (Last)		: Hong Chang Ho
'* 11. Comment		: 
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID = "b1b13mb1.asp"			'��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "b1b13mb2.asp"			'��: Save �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMPITEMBYPLANT_ID = "b1b11ma1"

Dim C_AltItemCd 
Dim C_ItemPopup 
Dim C_AltItemNm 
Dim C_AltItemSpec	'Added by SW 2003-01-07
Dim C_Priority  
Dim C_ValidFromDt 
Dim C_ValidToDt
Dim C_Seq		

Dim BaseDate
Dim StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop						' Popup
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

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
	lgSortKey = 1
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

Sub SetCookieVal()
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm") 
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm",""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm",""

End Sub


'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================

Sub SetPlant()
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			
			frm1.txtItemCd.focus
			Set gActiveElement = document.activeElement  
		Else
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement  
		End If	
	End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
     
   	Call InitSpreadPosVariables()

    With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021123", , Parent.gAllowDragDropSpread

		.ReDraw = False
	
		.MaxCols = C_Seq + 1
		.MaxRows = 0
    
		Call AppendNumberPlace("6","3","0")

		Call GetSpreadColumnPos("A")
       
		ggoSpread.SSSetEdit		C_AltItemCd,	"��üǰ��", 20,,,18,2
		ggoSpread.SSSetButton 	C_ItemPopup
		ggoSpread.SSSetEdit		C_AltItemNm,	"��üǰ���", 30,,,40  
		ggoSpread.SSSetEdit		C_AltItemSpec,	"��üǰ��԰�", 30,,,40  
		ggoSpread.SSSetFloat	C_Priority,		"�켱����",8,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,,"Z"
		ggoSpread.SSSetDate 	C_ValidFromDt,	"������", 12, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,	"������", 12, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_Seq,			"����", 6, 2
    
		Call ggoSpread.MakePairsColumn(C_AltItemCd, C_ItemPopup )
		Call ggoSpread.SSSetColHidden(C_Seq, C_Seq, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
		ggoSpread.SSSetSplit2(1)										'frozen ����߰� 
				
		Call SetSpreadLock 

		.ReDraw = True

    End With
    
End Sub


'========================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'======================================================================================== 
Sub InitSpreadPosVariables()

	C_AltItemCd   = 1
	C_ItemPopup   = 2
	C_AltItemNm   = 3
	C_AltItemSpec = 4
	C_Priority    = 5
	C_ValidFromDt = 6
	C_ValidToDt	  = 7
	C_Seq		  = 8
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

		C_AltItemCd   = iCurColumnPos(1)
		C_ItemPopup   = iCurColumnPos(2)
		C_AltItemNm   = iCurColumnPos(3)
		C_AltItemSpec = iCurColumnPos(4)
		C_Priority    = iCurColumnPos(5)
		C_ValidFromDt = iCurColumnPos(6)
		C_ValidToDt	  = iCurColumnPos(7)
		C_Seq		  = iCurColumnPos(8)
		
	End Select

End Sub



'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim IntRetCD

	'Call SetPopupMenuItemInf("1101111111")	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000110111")
	Else 	
		If frm1.vspdData.MaxRows = 0 Then 
			Call SetPopupMenuItemInf("1001111111")
		Else
			Call SetPopupMenuItemInf("1101111111") 
		End if			
	End If	
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	If frm1.vspdData.MaxRows = 0 Or Col < 0 Then
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
	
	'------ Developer Coding part (Start)
	'------ Developer Coding part (End)
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

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
  
     With frm1

    .vspdData.ReDraw = False
	
	ggoSpread.SSSetRequired		C_AltItemCd,	-1		
	ggoSpread.SpreadLock		C_AltItemNm,	-1, C_AltItemNm
	ggoSpread.SpreadLock		C_AltItemSpec,	-1, C_AltItemSpec
	ggoSpread.SSSetRequired		C_Priority,		-1
	ggoSpread.SpreadLock		C_ValidFromDt,	-1, C_ValidFromDt
	ggoSpread.SSSetRequired 	C_ValidToDt,	-1		
	ggoSpread.SSSetProtected	.vspdData.MaxCols, -1
	
	.vspdData.ReDraw = True
	
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
       
    With frm1
    
		.vspdData.ReDraw = False
	
		ggoSpread.SSSetRequired  C_AltItemCd,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AltItemNm,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AltItemSpec, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_Priority ,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_ValidFromDt, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_ValidToDt,	pvStartRow, pvEndRow  
		
		.vspdData.ReDraw = True
    
    End With
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

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function


'------------------------------------------  OpenConItemCd()  -------------------------------------------------
'	Name : OpenConItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemCd(ByVal strCode, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = strCode
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    arrField(2) = 3								' Field��(2) : "SPEC"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet, iPos)
	End If	
	
	If iPos = 0 Then	
		Call SetFocusToDocument("M")
		frm1.txtItemCd.focus
	Else
		Call SetActiveCell(frm1.vspdData,C_AltItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	

End Function

'==========================================  2.4.3 Set Return Value()  =============================================
'	Name : Set Return Value()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetItemPopup()  --------------------------------------------------
'	Name : SetItemPopup()
'	Description : OpenItemPopup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetItemInfo(Byval arrRet, ByVal iPos)
	With frm1
		If iPos = 0 Then	
			.txtItemCd.value = arrRet(0)
			.txtItemNm.value = arrRet(1)		
		ElseIf iPos = 1 Then
			.vspdData.Row = .vspdData.ActiveRow

			.vspdData.Col = C_AltItemCd
			.vspdData.Text = arrRet(0)
			
			.vspdData.Col = C_AltItemNm
			.vspdData.Text = arrRet(1)		
			
			.vspdData.Col = C_AltItemSpec
			.vspdData.Text = arrRet(2)		
			
			Call vspdData_Change(.vspdData.Col, .vspdData.Row)
		End If
	End With

End Function


'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)

    frm1.txtPlantCd.Value    = arrRet(0)		
    frm1.txtPlantNm.Value    = arrRet(1)
    
End Function

Function JumpItemByPlant()
	Dim IntRetCD
	
	ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900017",parent.VB_YES_NO,"X","X")            '��: Display Message(There is no changed data.)
        If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value  
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 
	WriteCookie "MainFormFlg", "ALTITEM"
	
	PgmJump(BIZ_PGM_JUMPITEMBYPLANT_ID)
	

End Function

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
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field

	Call InitSpreadSheet                                                    '��: Setup the Spread sheet
	Call InitVariables                                                      '��: Initializes local global variables
	Call SetCookieVal
	
	'----------  Coding part  -------------------------------------------------------------
	'Call SetToolbar("11001101000011")										'��: ��ư ���� ���� 
	Call SetToolbar("11000000000011")										'��: ��ư ���� ����	
   
	Call SetPlant	
		
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
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub




'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspddata_KeyPress(index , KeyAscii )
     
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()

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
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop)	Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	'----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_ItemPopup Then
        .Col = C_AltItemCd
        .Row = Row

        Call OpenConItemCd (.Text,1)
        
        Call SetActiveCell(frm1.vspdData,C_AltItemCd,Row,"M","X","X")
		Set gActiveElement = document.activeElement
     End If
    
    End With
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
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	  
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	Call ggoSpread.ClearSpreadData
    Call SetDefaultVal
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
    End If     												'��: Query db data

    FncQuery = True																'��: Processing is OK

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
    Dim IntRetCD 
    Dim lRow
    Dim starDate
    Dim finaDate
    
    FncSave = False																'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing
    On Error Resume Next														'��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    
    If Not ggoSpread.SSDefaultCheck Then              '��: Check required field(Multi area)
		Exit Function
    End If  
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If Not chkField(Document, "1") Then									'��: This function check indispensable field
			Exit Function
		End If			
    End If
    
    With frm1
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col= C_ValidFromDt
			starDate =  Trim(.vspdData.Text)
        
			.vspdData.Col= C_ValidToDt
			finaDate =  Trim(.vspdData.Text)
        
			If UNIConvDate(finaDate) < UNIConvDate(starDate) Then
				Call DisplayMsgBox("972002", "X", "������", "������")
				Call SheetFocus(lRow, C_ValidToDt)
				Exit Function
			End If
		Next
	End With
	
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     																				'��: Save db data
    
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
	frm1.vspdData.EditMode = True
	    
	frm1.vspdData.ReDraw = False    
	    
    ggoSpread.Source = frm1.vspdData	    
        
    ggoSpread.CopyRow   
    
    With frm1	
		
		.vspdData.Col = C_ValidFromDt
		.vspdData.Row = .vspdData.ActiveRow 
    
		.vspdData.Text = StartDate
    
		.vspdData.Col = C_ValidToDt
		.vspdData.Row = .vspdData.ActiveRow 
    
		.vspdData.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
    
		frm1.vspdData.ReDraw = True    
        
	    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow   
    
    End With
End Function


'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================

Function FncPaste() 
     ggoSpread.SpreadPaste
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt)  

    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If
	
    With frm1	
		
		.vspdData.ReDraw = False
		.vspdData.focus

	    ggoSpread.Source = .vspdData
        ggoSpread.InsertRow , iIntReqRows

		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)

		For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + iIntReqRows - 1
			.vspdData.Row = iIntCnt 
			.vspdData.Col = C_ValidFromDt
			.vspdData.Text = StartDate
    
			.vspdData.Col = C_ValidToDt
			.vspdData.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
		Next
		
		.vspdData.ReDraw = True
     
    End With    

    Set gActiveElement = document.activeElement 

	If Err.number = 0 Then
		FncInserRow = True
	End IF

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt

    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData 
	lDelRows = ggoSpread.DeleteRow
    
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
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
    
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()												'��: ���� ������ ���� ���� 
	
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'******************************************************************************************************

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
    
    LayerShowHide(1)
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows		
    Else   
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)		
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows	
    End If
   
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()				'��: ��ȸ ������ ������� 
	
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
	
	frm1.hPlantCd.value = Trim(frm1.txtPlantCd.value)
	frm1.hItemCd.value = Trim(frm1.txtItemCd.value)

End Function



'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
   	Dim strVal, strDel
   	Dim iColSep
   	Dim TmpBufferVal, TmpBufferDel
   	Dim iTotalStrVal, iTotalStrDel
   	Dim iValCnt, iDelCnt
	Dim starDate
	Dim finaDate
	
    DbSave = False                                                          '��: Processing is NG
    
    LayerShowHide(1)
		
    On Error Resume Next
                                                       '��: Protect system from crashing
	With frm1
		 .txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = Parent.gColSep
    lGrpCnt = 1
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag												'��: �ű� 
				
				strVal = ""
				
				strVal = strVal & "C" & iColSep & lRow & iColSep					'��: C=Create
                
                .vspdData.Col = C_AltItemCd
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                If UCase(Trim(.vspdData.Text)) = UCase(Trim(frm1.txtItemCd.value)) Then
					Call DisplayMsgBox("127421","X", "��üǰ��","ǰ��")
					Call LayerShowHide(0)
					Exit Function
				End If
					
                
                .vspdData.Col = C_Priority	
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_ValidFromDt	
                strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & iColSep

                .vspdData.Col = C_ValidToDt	
                strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gRowSep                
                
                ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = StrVal
                iValCnt = iValCnt + 1                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
            
				strVal = ""

				strVal = strVal & "U" & iColSep						'��: U=Update
				
                .vspdData.Col = C_Seq	
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_AltItemCd	

                If UCase(Trim(.vspdData.Text)) = UCase(Trim(frm1.txtItemCd.value)) Then
					Call DisplayMsgBox("127421","X", "��üǰ��","ǰ��")
					Call LayerShowHide(0)
					Exit Function
				End If
                
                strVal = strVal & Trim(.vspdData.Text) & iColSep       
               
                
                .vspdData.Col = C_Priority	
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_ValidFromDt	
                strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & iColSep
                
                .vspdData.Col = C_ValidToDt	
                strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gRowSep
                
				ReDim Preserve TmpBufferVal(iValCnt)
                
                TmpBufferVal(iValCnt) = StrVal
                iValCnt = iValCnt + 1                                                                                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag												'��: ���� 
            
				strDel = ""

				strDel = strDel & "D" & iColSep	
                .vspdData.Col = C_Seq
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBufferDel(iDelCnt)
                
                TmpBufferDel(iDelCnt) = StrDel
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

Function DbSaveOk()													'��: ���� ������ ���� ���� 

	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
	Call MainQuery()

End Function


Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��üǰ���</font></td>
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
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>									
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConItemCd frm1.txtItemCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>									
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
								<script language =javascript src='./js/b1b13ma1_I576177716_vspdData.js'></script>
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
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpItemByPlant">���庰ǰ���������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
