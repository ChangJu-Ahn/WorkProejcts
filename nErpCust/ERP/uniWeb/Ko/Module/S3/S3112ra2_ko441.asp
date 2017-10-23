
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p4111pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Production Order Reference ASP											*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/03/29																*
'*  8. Modified date(Last)  : 2002/12/10																*
'*  9. Modifier (First)     : Kim GyoungDon																*
'* 10. Modifier (Last)      : RYU SUNG WON																*
'* 11. Comment              :																			*
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)  
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin		                *
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->

<!--********************************************  1.1 Inc ����  *****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--============================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--============================================  1.1.2 ���� Include  ===================================
'=====================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
Const BIZ_PGM_QRY_ID = "S3112rb2_ko441.asp"			'��: �����Ͻ� ���� ASP�� 
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
Dim C_PLANT_CD		'PLANT
Dim C_PACK_LIST		'PACKING LIST
Dim C_ITEM_CD		'ǰ���ڵ�
Dim C_ITEM_NM		'ǰ���
Dim C_DN_TYPE		'����TYPE
Dim C_QTY			'����
Dim C_ISSUE_DT		'��������




	
'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
	
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
Dim arrReturn
Dim lgPlantCD
Dim strFromStatus
Dim strToStatus
Dim strThirdStatus
Dim IsOpenPop
Dim IsFormLoaded
Dim arrParent
Dim arrParam

	arrParent = window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(1)


top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable�� ����  ============================
'========================================================================================================
'----------------  ���� Global ������ ����  -------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++

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
	
	C_PLANT_CD	 = 1		'PLANT
	C_PACK_LIST	 = 2		'PACKING LIST
	C_ITEM_CD	 = 3		'ǰ���ڵ�
	C_ITEM_NM	 = 4		'ǰ���
	C_DN_TYPE	 = 5		'����TYPE
	C_QTY		 = 6		'����
	C_ISSUE_DT	 = 7		'��������

End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()

	Redim arrReturn(0,0)

    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1

	Self.Returnvalue = arrReturn
	
	txtBpCd.value			= arrParam(0)

End Function

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()

	Dim ArgArray						<%'Arguments�� �Ѱܹ��� Array%>
	'ArrParent = window.dialogArguments

	ArgArray  = ArrParent(1)
	




End Function

'==========================================   2.1.3 InitComboBox()  =====================================
'=	Name : InitComboBox()																				=
'=	Description : ComboBox�� Value�� Setting�Ѵ�.														=
'========================================================================================================
Sub InitComboBox()

    Dim iCodeArr 
    Dim iNameArr

  
   
	
	
 	 
    
End Sub

'==========================================  2.1.4 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display in Spread(s)
'========================================================================================================= 
Sub InitSpreadComboBox()
	Dim iCodeArr 
    Dim iNameArr

  
	
	
End Sub
'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData()
	Dim intRow
	Dim intIndex
	
	With vspdData
		
	End With
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================%>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE","PA") %>
	<% Call loadBNumericFormatA("Q", "P", "NOCOOKIE","PA") %>
End Sub
	
'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	
	vspdData.MaxCols = C_ISSUE_DT + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
		
        ggoSpread.SSSetEdit     C_PLANT_CD,			"����",				10,,, 5,2
        ggoSpread.SSSetEdit     C_PACK_LIST,		"PACKING LIST",		15,,, 15,2
        ggoSpread.SSSetEdit     C_ITEM_CD,			"ǰ���ڵ�",			12,,, 12,2
        ggoSpread.SSSetEdit     C_ITEM_NM,			"ǰ���",			18,,, 18,2
        ggoSpread.SSSetEdit     C_DN_TYPE,			"��������",         10,,, 10,2
		ggoSpread.SSSetFloat	C_QTY,				"����",				8,		popupparent.ggQtyNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	popupparent.gComNum1000, popupparent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit     C_ISSUE_DT,			"��������",			20,,, 20,2
		

		Call ggoSpread.SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols, True)
		vspdData.ReDraw = True
		vspdData.OperationMode = 5 
	Call SetSpreadLock()
End Sub
	
'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = vspdData
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
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		
			C_PLANT_CD	 = iCurColumnPos(1)		'PLANT
			C_PACK_LIST	 = iCurColumnPos(2)		'PACKING LIST
			C_ITEM_CD	 = iCurColumnPos(3)		'ǰ���ڵ�
			C_ITEM_NM	 = iCurColumnPos(4)		'ǰ���
			C_DN_TYPE	 = iCurColumnPos(5)		'����TYPE
			C_QTY		 = iCurColumnPos(6)		'����
			C_ISSUE_DT	 = iCurColumnPos(7)		'��������
		
            		
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
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call initData()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
 
    If OldLeft <> NewLeft Then Exit Sub
    
    if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then

		If lgStrPrevKeyIndex <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then Exit Sub
		End If
    End if    
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Sub txtBpCd_OnChange()

	Dim strSelect, strFrom, strWhere
	Dim IntRetCD
		Dim arrVal1, arrVal2
		Dim ii, jj

	lgBlnFlgChgValue = True

		If Trim(txtBpCd.value) <> "" Then
			strSelect = "bp_cd, bp_nm, bp_alias_nm "
			strFrom	  = " B_Biz_Partner "
			strWhere  = " bp_cd = " & FilterVar(LTrim(RTrim(txtbpCd.value)), "''", "S")
			
			 If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
				IntRetCD = DisplayMsgBox("211145","X","X","X")
				'frm1.txtCd.value = ""
                                'frm1.txtCostNM.value = ""
				'frm1.txtWcCd.value   = ""
                                'frm1.txtWcNm.value   = ""
				txtBpCd.focus	
			 Else
				arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
				
				jj = Ubound(arrVal1,1)

				For ii = 0 to jj - 1
					arrVal2 = Split(arrVal1(ii), chr(11))
					txtBpCd.value = Trim(arrVal2(1))
					txtBpNm.value = Trim(arrVal2(2))
					hBpAliasNm.value = Trim(arrVal2(3))
				Next
			End IF
	End If 		
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow

	If vspdData.MaxRows = 0 then
		Exit Function
	end if

	
		intInsRow = 0
		Redim arrReturn(ggoSpread.Source.SelBlockRow2 - ggoSpread.Source.SelBlockRow , vspdData.MaxCols-2)
		For intRowCnt = ggoSpread.Source.SelBlockRow To ggoSpread.Source.SelBlockRow2
			vspdData.Row = intRowCnt
			For intColCnt = 0 To vspdData.MaxCols - 2
				vspdData.Col = intColCnt + 1
				arrReturn(intInsRow, intColCnt) = vspdData.Text
			Next
			intInsRow = intInsRow + 1
	    Next 
	
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
	
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


Sub vspdData_KeyPress(keyAscii)
	If keyAscii=13 and vspdData.ActiveRow > 0 Then
 		Call OkClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub	



'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================


'

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    		
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	Call InitVariables
	Call txtBpCd_OnChange											'��: Initializes local global variables
	Call InitSpreadSheet()
		
	IsFormLoaded = true											'After Loading the Form, the OrderStatus Variables can be Changed.
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
Function FncQuery()
    FncQuery = False

    ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData
    'Call InitVariables
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

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then Exit Function
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

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
    Call GetSpreadColumnPos("A")
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


'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

  
    lgKeyStream       = Trim(hBpAliasNm.Value) & PopupParent.gColSep '0       
    lgKeyStream       = lgKeyStream & Trim(txtItemCd.Value) & PopupParent.gColSep '1                                'You Must append one character(parent.gColSep)
	lgKeyStream       = lgKeyStream & Trim(txtPACK_LIST.Value) & PopupParent.gColSep '1                                'You Must append one character(parent.gColSep)
       
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
	Dim strVal
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
		Exit Function
	End If
	call MakeKeyStream("X")

		
		strVal = BIZ_PGM_QRY_ID & "?txtMode="            & PopupParent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream  
        strVal = strVal & "&txtMaxRows="         & vspdData.MaxRows
		strVal = strVal & "&lgStrPrevKeyIndex="  & lgStrPrevKeyIndex  

	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                          
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()		
											'��: ��ȸ ������ ������� 
	If lgIntFlgMode = PopupParent.OPMD_CMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = PopupParent.OPMD_UMODE	
	Call InitData()
	
    vspddata.Focus												'��: Indicates that current mode is Update mode
End Function


    
'------------------------------------------  OpenCondPlant()  ---------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------- 
Function OpenConBP()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim activateField
	
	If IsOpenPop = True Or UCase(txtBPCd.className) = UCase(popupparent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ŷ�ó"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_PARTNER"						' TABLE ��Ī 
	arrParam(2) = Trim(txtBPCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "�ŷ�ó"					' TextBox ��Ī 
	
    arrField(0) = "BP_CD"					' Field��(0)
    arrField(1) = "BP_NM"
    arrField(2)	= "BP_ALIAS_NM"				' Field��(1)
    
    arrHeader(0) = "�ŷ�ó�ڵ�"					' Header��(0)
    arrHeader(1) = "�ŷ�ó��"					' Header��(1)
    arrHeader(2) = "MES(�ŷ�ó��)"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConBP(arrRet, 0)
	End If
		
	Call SetFocusToDocument("M")
	txtBPCd.focus
	
End Function


'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConBP(byval arrRet, byval iPos)
	
		
	txtBPCd.Value    	= arrRet(0)		
	txtBPNm.Value    	= arrRet(1)
	hBpAliasNm.Value	= arrRet(2)
		

End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPACK()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim activateField
	
	If IsOpenPop = True Or UCase(txtPACK_LIST.className) = UCase(popupparent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "PACKING LIST"				' �˾� ��Ī 
	arrParam(1) = "T_IF_RCV_VIRTURE_OUT_KO441"						' TABLE ��Ī 
	arrParam(2) = Trim(txtPACK_LIST.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = " ERP_APPLY_FLAG = 'N' "							' Where Condition
	arrParam(5) = "PACKING LIST"					' TextBox ��Ī 
	
    arrField(0) = "PACK_LIST"					' Field��(0)
    'arrField(1) = "BP_NM"					' Field��(1)
    
    arrHeader(0) = "PACKING LIST"					' Header��(0)
    'arrHeader(1) = "�ŷ�ó��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPACK(arrRet, 0)
	End If
	
	Call SetFocusToDocument("M")
	txtPACK_LIST.focus
	
End Function


'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPACK(byval arrRet, byval iPos)
	
	txtPACK_LIST.Value    = arrRet(0)		
			'.txtBPNm.Value    = arrRet(1)
	
End Function

Function OpenITEM()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim activateField
	
	If IsOpenPop = True Or UCase(txtITEMCD.className) = UCase(popupparent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��"				' �˾� ��Ī 
	arrParam(1) = "B_ITEM"						' TABLE ��Ī 
	arrParam(2) = Trim(txtITEMCD.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "ǰ��"					' TextBox ��Ī 
	
    arrField(0) = "ITEM_CD"					' Field��(0)
    arrField(1) = "ITEM_NM"					' Field��(1)
    
    arrHeader(0) = "ǰ��"					' Header��(0)
    arrHeader(1) = "ǰ���"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConITEM(arrRet, 0)
	End If
	
	Call SetFocusToDocument("M")
	txtITEMCD.focus
	
End Function


'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConITEM(byval arrRet, byval iPos)
	
	txtITEMCD.Value    = arrRet(0)		
	txtITEMNM.Value    = arrRet(1)
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
						
						
						<TR>  <TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtBPCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="�ŷ�ó�ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBPCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConBP()">
										<INPUT TYPE=TEXT NAME="txtBPNm" SIZE=20 tag="24"></TD>	
				     		   <TD CLASS=TD5 NOWRAP>ǰ��</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenITEM()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
				     	 
				     		</TD>
				     	
						</TR>		
				        <TR>		
				        <TD  CLASS="TD5" nowrap>PACKING LIST</TD>
				     	<TD CLASS="TD6" nowrap><INPUT NAME="txtPACK_LIST" ALT="PACKING LIST" TYPE="Text" SiZE=16 MAXLENGTH=20  tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPACK()">
				     	</TD>
				     	 
					   </TR>
					   </TABLE>
						</FIELDSET>
					</TD>
				</TR>
	<TR>
		<TD HEIGHT=100%>
		<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" ID=vspdData TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBpAliasNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrderType" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hToStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
