
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: Reference Popup Component List											*
'*  3. Program ID			: p4212ra1																	*
'*  4. Program Name			: �����Ȳ����																*
'*  5. Program Desc			: Reference Popup															*
'*  7. Modified date(First)	: 2000/04/06																*
'*  8. Modified date(Last)	: 2002/12/20																*
'*  9. Modifier (First)    	: Kim, Gyoung-Don															*
'* 10. Modifier (Last)		: Ryu Sung Won																*
'* 11. Comment 				:																			*
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)   
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin                        *
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--'####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->
<!--'********************************************  1.1 Inc ����  ****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 ���� Include  ==================================
'=====================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p4212rb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p4212rb2.asp"								'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ====================================
'==========================================================================================================

' Grid 1(vspdData1) - Operation
Dim C_SlCd
Dim C_SlNm
Dim C_GoodOnHandQty
Dim C_SchdRcptQty
Dim C_SchdIssueQty
Dim C_AvailQty

' Grid 2(vspdData2) - Operation
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_GoodOnHandQty1
Dim C_BlockIndicator

'==========================================  1.2.2 Global ���� ����  ==================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->
Dim lgIntPrevKey
Dim lgStrPrevKey2
Dim lgCurrRow
Dim IsOpenPop 

Dim lgPlantCD
Dim lgItemCD
Dim lgItemNm
Dim lgSlCD
Dim lgSlNm

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim lgOldRow

'*********************************************  1.3 �� �� �� ��  ****************************************
'*	����: Constant�� �ݵ�� �빮�� ǥ��.																*
'********************************************************************************************************
Dim arrParent
Dim arrParam					
		
'------ Set Parameters from Parent ASP ------
arrParent	= window.dialogArguments
Set PopupParent = arrParent(0)
lgPlantCD	= arrParent(1)
lgItemCD	= arrParent(2)
lgItemNm	= arrParent(3)
lgSlCD		= arrParent(4)
lgSlNm		= arrParent(5)
	
'top.document.title = "�����Ȳ"
top.document.title = PopupParent.gActivePRAspName

'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
			C_SlCd				= 1
			C_SlNm				= 2
			C_GoodOnHandQty		= 3
			C_SchdRcptQty		= 4
			C_SchdIssueQty		= 5
			C_AvailQty			= 6

		Case "B"
			C_TrackingNo		= 1
			C_LotNo				= 2
			C_LotSubNo			= 3
			C_GoodOnHandQty1	= 4
			C_BlockIndicator	= 5
	End Select			
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0							<%'��: Initializes Group View Size%>
	lgStrPrevKey = ""                           'initializes Previous Key		
	Self.Returnvalue = Array("")
End Function

'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================
Sub SetDefaultVal()

End Sub

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()
		txtPlantCd.value = lgPlantCD
		txtItemCd.value  = lgItemCD
		txtItemNm.value  = lgItemNm
		txtSlCd.value    = lgSlCD
		txtSlNm.value    = lgSlNm
End Function

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
			'------------------------------------------
			' Grid 1 - Operation Spread Setting
			'------------------------------------------
			Call InitSpreadPosVariables(pvSpdNo)
			ggoSpread.Source = vspdData1
			ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

			With vspdData1 
			.ReDraw = false
			.MaxCols = C_AvailQty + 1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			
			Call GetSpreadColumnPos(pvSpdNo)
			
			ggoSpread.SSSetEdit		C_SlCd,			"â��", 10
			ggoSpread.SSSetEdit		C_SlNm,			"â���", 20
			ggoSpread.SSSetFloat	C_GoodOnHandQty,"��ǰ���",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchdRcptQty,	"�԰���",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchdIssueQty, "�����",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_AvailQty,		"�������",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
			
			Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			.ReDraw = true
			End With
	
	
	'------------------------------------------
	' Grid 2 - Component Spread Setting
	'------------------------------------------
		Case "B"
			'------------------------------------------
			' Grid 2 - Component Spread Setting
			'------------------------------------------
			Call InitSpreadPosVariables("B")
			ggoSpread.Source = vspdData2
			ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread
	
			With vspdData2
			.ReDraw = false		
			.MaxCols = C_BlockIndicator + 1													'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0

			Call GetSpreadColumnPos("B")
		
			ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25
			ggoSpread.SSSetEdit		C_LotNo,		"Lot No.", 18
			ggoSpread.SSSetEdit		C_LotSubNo,		"Lot Sub No.", 10
			ggoSpread.SSSetFloat	C_GoodOnHandQty1, "��ǰ���",20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"						
			ggoSpread.SSSetEdit		C_BlockIndicator, "Block", 15
	
			Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			.ReDraw = true
			End With
	End Select
	
	Call SetSpreadLock()
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    '--------------------------------
    'Grid 1
    '--------------------------------
    ggoSpread.Source = vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
    
    '--------------------------------
    'Grid 2
    '--------------------------------
    ggoSpread.Source = vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
            ggoSpread.Source = vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_SlCd				= iCurColumnPos(1)
			C_SlNm				= iCurColumnPos(2)
			C_GoodOnHandQty		= iCurColumnPos(3)
			C_SchdRcptQty		= iCurColumnPos(4)
			C_SchdIssueQty		= iCurColumnPos(5)
			C_AvailQty			= iCurColumnPos(6)
		Case "B"
			ggoSpread.Source = vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_TrackingNo		= iCurColumnPos(1)
			C_LotNo				= iCurColumnPos(2)
			C_LotSubNo			= iCurColumnPos(3)
			C_GoodOnHandQty1	= iCurColumnPos(4)
			C_BlockIndicator	= iCurColumnPos(5)
    End Select
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet(gActiveSpdSheet.Id)
	
    If gActiveSpdSheet.Id = "A" Then
		ggoSpread.Source = vspdData1
	Else
		ggoSpread.Source = vspdData2
	End If
	
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub vspdData1_KeyPress(keyAscii)
	If keyAscii=27 Then
 		Call CancelClick()
		Exit Sub
	End If
End Sub	

Sub vspdData2_KeyPress(keyAscii)
	If keyAscii=27 Then
 		Call CancelClick()
		Exit Sub
	End If
End Sub	

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	If txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field��(0)
	arrField(1) = 2 '"ITEM_NM"					' Field��(1)
    
    iCalledAspName = AskPRAspName("b1b11pa3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtItemCd.focus

End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = Trim(txtSLCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(txtPlantCd.value), "''", "S") 	' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
	
    arrField(0) = "SL_CD"													' Field��(0)
    arrField(1) = "SL_NM"													' Field��(1)
    
    arrHeader(0) = "â��"												' Header��(0)
    arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	txtSLCd.focus
	
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet)
	txtItemCd.Value    = arrRet(0)		
	txtItemNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSLCd(byval arrRet)
	txtSLCd.Value    = arrRet(0)		
	txtSLNm.Value    = arrRet(1)		
End Function

'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'��: Load table , B_numeric_format			
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call SetDefaultVal
	Call InitVariables											'��: Initializes local global variables
	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")
	Call InitSetting()
	Call ggoOper.LockField(Document, "N")									'��: This function lock the suitable field
	
	Call FncQuery()

End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery

	FncQuery = False
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	vspddata1.MaxRows = 0
	vspddata2.MaxRows = 0
	If DbQuery = False Then	
		Exit Function
	End If
	FncQuery = False
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

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = vspdData1
	Call SetPopupMenuItemInf("0000111111")
	
	If vspdData1.MaxRows <= 0 Then Exit Sub
	
	If Row <= 0 Then
        ggoSpread.Source = vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
	
	If lgOldRow <> Row Then
		
		vspdData1.Col = 1
		vspdData1.Row = row
		
		lgOldRow = Row
		
		vspdData2.MaxRows = 0
	  	
		If DbDtlQuery = False Then	
			Exit Sub
		End If	
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SP2C"					'SpreadSheet ������ vspdData�ϰ�� 
	Set gActiveSpdSheet = vspdData2
	Call SetPopupMenuItemInf("0000111111")
	
    If vspdData2.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
    Err.Clear												'��: Protect system from crashing
	    
    DbQuery = False											'��: Processing is NG
	    
    Call LayerShowHide(1)
	    
    Dim strVal

    strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & PopupParent.UID_M0001		'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & txtPlantCd.value 	'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & txtItemCd.value 		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtSlCd=" & txtSLCd.value 			'��: ��ȸ ���� ����Ÿ 
	    
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    

    Call RunMyBizASP(MyBizASP, strVal)						'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                          '��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	Dim LngRow
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    
    With vspdData1
		.ReDraw = False
		If .MaxRows > 0 Then
			For LngRow = 1 To .MaxRows
				.Row = LngRow
				.Col = C_AvailQty
				If uniCDbl(.Text) < 0 Then
					.ForeColor = vbRed
					.Col = C_SlCd
					.ForeColor = vbRed
				End If
			Next
		End If
		.ReDraw = True
	End With
    
    Call SetActiveCell(vspdData1,1,1,"P","X","X")
	Set gActiveElement = document.activeElement
	Call DbDtlQuery
	vspdData1.Focus

End Function

Function DbQueryNotOk()														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
	vspdData1.Focus

End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal
   
		DbDtlQuery = False   
    
		vspdData1.Row = vspdData1.ActiveRow

		Call LayerShowHide(1)

			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtPlantCd=" & txtPlantCd.value
			strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)
			vspdData1.Col = C_SLCd
			strVal = strVal & "&txtSlCd=" & Trim(vspdData1.Text)
			strVal = strVal & "&txtMaxRows=" & vspdData2.MaxRows
			
		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 

End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5 NOWRAP>����</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14xxxU" ALT="����">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
						<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>ǰ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>â��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="â��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=20 tag="14" ALT="â���"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>�԰�</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=40 tag="14" ALT="�԰�">&nbsp;</TD>
						<TD CLASS=TD5 NOWRAP>�������</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p4212ra1_I379220533_txtSafetyStock.js'></script></TD>
					</TR>	
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TR HEIGHT="40%">
			<TD WIDTH="100%" colspan=4>
				<script language =javascript src='./js/p4212ra1_A_vspdData1.js'></script>
			</TD>
		</TR>
		<TR HEIGHT="60%">
			<TD WIDTH="100%" colspan=4>
				<script language =javascript src='./js/p4212ra1_B_vspdData2.js'></script>
			</TD>
		</TR>	
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hSlCd" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
