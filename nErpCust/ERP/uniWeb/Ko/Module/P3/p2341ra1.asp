
<%@ LANGUAGE="VBSCRIPT" %>
<%'********************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: Reference Popup Pegging List											*
'*  3. Program ID			: p2341ra1																	*
'*  4. Program Name			: Pegging ����																*
'*  5. Program Desc			: Reference Popup															*
'*  7. Modified date(First)	: 2003/11/04																*
'*  8. Modified date(Last)	: 																*
'*  9. Modifier (First)    	: Chen, Jae Hyun											*
'* 10. Modifier (Last)		: 																*
'* 11. Comment 				:																			*
'* 12. History              : 
'*                          :                   *
'********************************************************************************************************%>

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
<SCRIPT LANGUAGE="VBScript"	  SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "p2341rb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY2_ID	= "p2341rb2.asp"								'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
' Grid 1(vspdData) - MRP info
Dim C_RqrDt
Dim C_ParentItemCd
Dim C_ParentItemNm
Dim C_Spec
Dim C_MPSQty
Dim C_SchdIssQty
Dim C_RqrQty
Dim C_TotalRqrQty
Dim C_SchdRcptQty
Dim C_PrevAvailQty
Dim C_PlanQty

'==========================================  1.2.2 Global ���� ����  ==================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'======================================================================================================
Const C_Sep  = "/"
Const C_PROD  = "PROD"
Const C_MATL  = "MATL"
Const C_PHANTOM ="PHANTOM"
Const C_ASSEMBLY = "ASSEMBLY"
Const C_SUBCON  = "SUBCON"

Const C_IMG_PROD = "../../../CShared/image/product.gif"
Const C_IMG_MATL = "../../../CShared/image/material.gif"
Const C_IMG_PHANTOM = "../../../CShared/image/phantom.gif"
Const C_IMG_ASSEMBLY = "../../../CShared/image/Assembly.gif"
Const C_IMG_SUBCON = "../../../CShared/image/subcon.gif"

Const tvwChild = 4

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgBlnBizLoadMenu
Dim lgSelNode
Dim lgStrColorFlag

'Dim lgStrPrevKey

Dim lgPlantCD
Dim lgItemCd
Dim lgBomNo

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
Dim IsOpenPop			'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim lgOldRow

'*********************************************  1.3 �� �� �� ��  ****************************************
'*	����: Constant�� �ݵ�� �빮�� ǥ��.																*
'********************************************************************************************************
Dim arrParent
Dim arrParam					
		
'------ Set Parameters from Parent ASP ------
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
lgPlantCD		= arrParent(1)
lgItemCd		= arrParent(2)
lgBomNo			= "1"
	
top.document.title = PopupParent.gActivePRAspName

Dim BaseDate

BaseDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

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
Sub InitSpreadPosVariables()
	C_RqrDt				= 1
	C_ParentItemCd		= 2
	C_ParentItemNm		= 3
	C_Spec              = 4
	C_MPSQty			= 5
	C_SchdIssQty		= 6
	C_RqrQty			= 7
	C_TotalRqrQty		= 8
	C_SchdRcptQty		= 9
	C_PrevAvailQty		= 10
	C_PlanQty			= 11
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgStrPrevKey = ""
	Self.Returnvalue = Array("")
	lgSelNode = ""
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
	txtPlantCd.value = lgPlantCd
	txtItemCd.value = lgItemCd
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

	With vspdData
		.ReDraw = false
		.MaxCols = C_PlanQty +1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0

		Call GetSpreadColumnPos()
		
		ggoSpread.SSSetDate 	C_RqrDt,		 		"����"	,		11, 2, PopupParent.gDateFormat    
		ggoSpread.SSSetEdit		C_ParentItemCd,			"��ǰ��",		18
		ggoSpread.SSSetEdit		C_ParentItemNm,			"��ǰ���",		25
		ggoSpread.SSSetEdit		C_Spec,			        "�԰�",		    25
		ggoSpread.SSSetEdit		C_MPSQty,				"MPS����"		,15, 1
		ggoSpread.SSSetEdit		C_SchdIssQty,			"�����"		,15, 1
		ggoSpread.SSSetEdit		C_RqrQty,				"�ҿ䷮"		,15, 1
		ggoSpread.SSSetEdit		C_TotalRqrQty,			"�Ѽҿ䷮"		,15, 1
		ggoSpread.SSSetEdit		C_SchdRcptQty,			"�԰���"		,15, 1
		ggoSpread.SSSetEdit		C_PrevAvailQty,			"�������"		,15, 1
		ggoSpread.SSSetEdit		C_PlanQty,				"��ȹ����"		,15, 1
		

		Call ggoSpread.SSSetColHidden(C_ParentItemNm,C_ParentItemNm, True)
		Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols, True)
		
		ggoSpread.SSSetSplit2(2)
		.ReDraw = true
	End With
	    
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
    vspdData.ReDraw = False
    ggoSpread.Source = vspdData
	ggoSpread.SpreadLock 1 , -1	
	vspdData.ReDraw = True
    
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos()
	
	Dim iCurColumnPos
	
    ggoSpread.Source = vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
   	C_RqrDt				= iCurColumnPos(1)
	C_ParentItemCd		= iCurColumnPos(2)
	C_ParentItemNm		= iCurColumnPos(3)
	C_Spec              = iCurColumnPos(4)
	C_MPSQty			= iCurColumnPos(5)
	C_SchdIssQty		= iCurColumnPos(6)
	C_RqrQty			= iCurColumnPos(7)
	C_TotalRqrQty		= iCurColumnPos(8)
	C_SchdRcptQty		= iCurColumnPos(9)
	C_PrevAvailQty		= iCurColumnPos(10)
	C_PlanQty			= iCurColumnPos(11)

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
	Call InitSpreadSheet()

	ggoSpread.Source = vspdData
	Call ggoSpread.ReOrderingSpreadData()

End Sub

'========================== 2.2.6 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitComboBox()
    
End Sub


'=========================================  2.3.2 CancelClick()  ========================================
' Name : CancelClick()
' Description : Return Array to Opener Window for Cancel button click
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"					'SpreadSheet ������ vspdData�ϰ�� 
	Set gActiveSpdSheet = vspdData
	Call SetPopupMenuItemInf("0000111111")
	
    If vspdData.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = vspdData
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos()
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub



'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

End Sub

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

Sub vspdData_KeyPress(keyAscii)
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
	arrParam(1) = strCode					' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field��(0)
	arrField(1) = 2 '"ITEM_NM"					' Field��(1)
    
    iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
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

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet)
	txtItemCd.Value    = arrRet(0)		
	txtItemNm.Value    = arrRet(1)		
End Function

'========================================================================================
' Function Name : InitTreeImage
' Function Desc : �̹��� �ʱ�ȭ 
'========================================================================================
Function InitTreeImage()
	Dim NodX, lHwnd
	
	uniTree1.SetAddImageCount = 4
	uniTree1.Indentation = "200"									' �� ���� 
	uniTree1.AddImage C_IMG_PROD, C_PROD, 0						'��: TreeView�� ���� �̹��� ���� 
	uniTree1.AddImage C_IMG_MATL, C_MATL, 0
	uniTree1.AddImage C_IMG_ASSEMBLY, C_ASSEMBLY, 0				'��: TreeView�� ���� �̹��� ���� 
	uniTree1.AddImage C_IMG_PHANTOM, C_PHANTOM, 0
	uniTree1.AddImage C_IMG_SUBCON, C_SUBCON, 0

	uniTree1.OLEDragMode = 0										'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
	uniTree1.OLEDropMode = 0

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
	Call SetDefaultVal
	Call InitVariables											'��: Initializes local global variables
	Call ggoOper.LockField(Document, "N")						'��: This function lock the suitable field
	Call InitTreeImage()	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub


'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node Click�� Look Up Call
'==========================================================================================
Sub uniTree1_NodeClick(ByVal Node)
    Dim NodX
    
	Dim iPos1
	
	Dim txtItemCd

	Err.Clear                                                               '��: Protect system from crashing
	
    Set NodX = uniTree1.SelectedItem
    
    If Not NodX Is Nothing Then				' ���õ� ������ ������ 

		'-------------------------------------
		'If Same Node Clicked, Exit
		'---------------------------------------
			
		If NodX.Key = lgSelNode Then
			Set NodX = Nothing
			Exit Sub
		Else
			lgSelNode = NodX.Key
		End If
		
		iPos1 = InStr(NodX.Text, "    (")   
		txtItemCd = Trim(Left(NodX.Text,iPos1-1))
		
		vspdData.MaxRows = 0
		
		If DbDtlQuery(txtItemCd) = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
	End If	
		
    Set NodX = Nothing
    
End Sub


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

'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
        
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

	'-----------------------
    'Erase contents area
    '----------------------- 

	uniTree1.Nodes.Clear		
									'��: Tree View Content	 
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
            
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
    End If     																		'��: Query db data
       
    FncQuery = True																'��: Processing is OK
        
End Function

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************

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
	
    Dim strVal
    
    Err.Clear															'��: Protect system from crashing
    
    DbQuery = False														'��: Processing is NG
   
	LayerShowHide(1)
		
    '----------------------------------------------
    '- Call Query ASP
    '----------------------------------------------
    
    'frm1.txtUpdtUserId.value= parent.gUsrID
    
    strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(txtPlantCd.value)					'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)						'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBaseDt=" & BaseDate		
    							
    If rdoFlg(0).checked Then
		strVal = strVal & "&rdoSrchType=4"
	Else
		strVal = strVal & "&rdoSrchType=2"	
	End If
	
    If rdoFromFlg(0).checked Then
		strVal = strVal & "&rdoSrchType1=1"
	Else
		strVal = strVal & "&rdoSrchType1=2"	
	End If											'������ 

    strVal = strVal & "&txtBomNo="	& lgBOMNo									'BOM No.
   
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

    DbQuery = True														'��: Processing is NG
                                        '��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()											'��: ��ȸ ������ ������� 
	
	Dim NodX
	Dim iPos1
    Dim strItemCd
    
    '-----------------------
    'Reset variables area
    '-----------------------	
	Set NodX = uniTree1
	NodX.SetFocus
	
	If Not NodX.selectedItem Is Nothing Then				' ���õ� ������ ������ 
	
		If NodX.SelectedItem.Key = lgSelNode Then
			Set NodX = Nothing
			Exit Function
		Else
			lgSelNode = NodX.SelectedItem.Key
		End If
		
		iPos1 = InStr(NodX.SelectedItem.Text, "    (")   
		strItemCd = Trim(Left(NodX.SelectedItem.Text,iPos1-1))
		
	End If	
	
	Set NodX = Nothing
	Set gActiveElement = document.activeElement
	Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    
	vspdData.MaxRows = 0
	
	If DbDtlQuery(strItemCd) = False Then	
		Call RestoreToolBar()
		Exit Function
	End If	
	lgOldRow = 1	
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByVal strItemCd) 

	Dim strVal
    
    DbDtlQuery = False    
    
    Err.Clear
    
	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & PopupParent.UID_M0001
	strVal = strVal & "&txtPlantCd=" & Trim(txtPlantCd.value)
	strVal = strVal & "&txtItemCd=" & Trim(strItemCd)
	strVal = strVal & "&txtMaxRows=" & vspdData.MaxRows
		
	Call RunMyBizASP(MyBizASP, strVal)    
    
    DbDtlQuery = True

End Function


Function DbDtlQueryOk()													'��: ��ȸ ������ ������� 
	'--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call SetQuerySpreadColor
	'--------- Developer Coding Part (End) ----------------------------------------------------------
    Set gActiveElement = document.ActiveElement  
End Function

Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	
	iArrColor1 = Split(lgStrColorFlag,PopUpParent.gRowSep)
	
	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),PopUpParent.gColSep)
		
		vspdData.Col = -1
		vspdData.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				vspdData.BackColor = RGB(204,255,153) '���� 
			Case "2"
				vspdData.BackColor = RGB(176,234,244) '�ϴû� 
			Case "3"
				vspdData.BackColor = RGB(224,206,244) '������ 
			Case "4"  
				vspdData.BackColor = RGB(251,226,153) '����Ȳ 
			Case "5" 
				vspdData.BackColor = RGB(255,255,153) '����� 
		End Select
	Next

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=10>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>	
					<TR>
						<TD CLASS=TD5 NOWRAP>����</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14xxxU" ALT="����">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>ǰ��</TD>	
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoFromFlg" ID="rdoFromFlg1" CLASS="RADIO" tag="1X"CHECKED><LABEL FOR="rdoFromFlg1">MRP</LABEL>
										     <INPUT TYPE="RADIO" NAME="rdoFromFlg" ID="rdoFromFlg2" CLASS="RADIO" tag="1X"><LABEL FOR="rdoFromFlg2">BOM</LABEL></TD>
						<TD CLASS=TD5 NOWRAP>�������</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg1" CLASS="RADIO" tag="1X"CHECKED><LABEL FOR="rdoFlg1">������</LABEL>
										     <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg2" CLASS="RADIO" tag="1X"><LABEL FOR="rdoFlg2">������</LABEL></TD>										     
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=* WIDTH=25%>
						<script language =javascript src='./js/p2341ra1_uniTree1_N448168486.js'></script>
					</TD>
					<TD HEIGHT=* WIDTH=75% VAlign=Top>
						<TABLE CLASS="BasicTB" CELLSPACING=0>
							<TR>
								<TD WIDTH="100%">
									<script language =javascript src='./js/p2341ra1_I173195981_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>							
	</TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>