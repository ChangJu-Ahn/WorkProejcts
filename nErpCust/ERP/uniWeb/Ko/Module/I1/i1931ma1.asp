<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : BackLog �����(site)
'*  3. Program ID           : i1931ma1_ko119
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2006/04/11
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : HJO
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'#######################################################################################################
'												1. �� �� �� 
'#######################################################################################################
-->
<!--
'******************************************  1.1 Inc ����   ********************************************
'	���: Inc. Include
'*******************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

' Condition���� Default ��ȸ ��¥ 
Dim iDBSYSDate
Dim LocSvrDate
Dim StartDate
Dim EndDate

	iDBSYSDate = "<%=GetSvrDate%>"		
	LocSvrDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	StartDate = UNIDateAdd("D",-7,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
	EndDate = UNIDateAdd("D", 7,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()     
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I","*","NOCOOKIE","MA") %>

End Sub

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'Grid 1 - Order Header
Const BIZ_PGM_QRY1_ID	= "i1931mb1.asp"						'��: Head Query �����Ͻ� ���� ASP�� 
'Post Production Results
Const BIZ_PGM_SAVE_ID	= "i1931mb2.asp"						
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Order Header
Dim C_Chk
Dim C_ProdtDt
Dim C_ProdtOrderNo			
Dim C_OprNo					
Dim C_ItemCd				
Dim C_ItemNm				
Dim C_Spec					
Dim C_IssueQty	
Dim C_GoodQty				
Dim C_BasicUnit					
Dim C_SlCd				
Dim C_SlNm
	
Dim C_TrackingNo			
Dim C_ReqNo	
Dim C_ResvSeq
Dim C_ResultSeq
Dim C_DocumentNo
Dim C_Status
Dim C_StatusNm			
Dim C_Error
Dim C_PlantCd
Dim C_DocumentYear
Dim C_WcCd
Dim C_LotNo
Dim C_LotSubNo
Dim C_CostCd
Dim C_SchdQty

Dim C_OriginQty
Dim C_Remark
			


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntFlgMode								'Variable is for Operation Status
Dim lgIntPrevKey
Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgLngCurRows
Dim lgCurrRow
Dim lgCheckall			
Dim lgBlnFlgChgValue
'==========================================  1.2.3 Global Variable�� ����  ==================================
'============================================================================================================
'----------------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2
'++++++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

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
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""							'initializes Previous Key
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgOldRow = 0
	lgSortKey1   = 1
	lgSortKey2   = 1
	
	lgCheckall=0	
	lgBlnFlgChgValue = False
	frm1.btnRun.value = "��ü����"
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
    frm1.txtProdFromDt.text = StartDate
    frm1.txtProdToDt.text   = EndDate
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(pvSpdNo)

    Call InitSpreadPosVariables(pvSpdNo)
    
    Call AppendNumberPlace("6", "3", "0")
    Call AppendNumberPlace("7", "5", "0")
    
	If pvSpdNo = "A"  Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
	
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20050920", ,Parent.gAllowDragDropSpread
    
			.ReDraw = false
    
			.MaxCols = C_OriginQty + 1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ    
			.MaxRows = 0
    
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetCheck 		C_Chk,			"����", 8, ,,True, -1
			ggoSpread.SSSetDate 		C_ProdtDt,			"������", 11, 2, parent.gDateFormat     
			ggoSpread.SSSetEdit			C_ProdtOrderNo,			"����������ȣ", 18
			ggoSpread.SSSetEdit			C_OprNo,				"����", 6
			ggoSpread.SSSetEdit			C_ItemCd,				"ǰ��", 18
			ggoSpread.SSSetEdit			C_ItemNm,				"ǰ���", 20
			ggoSpread.SSSetEdit			C_Spec,			"�԰�", 20
			ggoSpread.SSSetFloat		C_IssueQty,		"��������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit			C_Remark,		"���", 30,,,50
			ggoSpread.SSSetFloat		C_GoodQty,		"������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit			C_BasicUnit,		"����", 8	
			
			ggoSpread.SSSetEdit			C_SlCd,					"â��", 10
			ggoSpread.SSSetEdit			C_SlNm,					"â���", 20
			ggoSpread.SSSetEdit			C_TrackingNo,			"Tracking No.", 25
			ggoSpread.SSSetEdit 		C_ReqNo,				"��û��ȣ",	15
			ggoSpread.SSSetEdit 		C_ResvSeq,				"�����Ϸù�ȣ",	15
			ggoSpread.SSSetEdit 		C_ResultSeq,			"�����������",	15
			ggoSpread.SSSetEdit 		C_DocumentNo,		"���ҹ�ȣ",	15
			ggoSpread.SSSetEdit 		C_Status,				"����",	10
			ggoSpread.SSSetEdit			C_StatusNm,				"����", 10
			ggoSpread.SSSetEdit 		C_Error,					"��������",	50
			ggoSpread.SSSetEdit 		C_PlantCd,				"����",	4
			ggoSpread.SSSetEdit 		C_DocumentYear,		"���ҳ⵵",4
			ggoSpread.SSSetEdit 		C_WcCd,			"�۾���",4
			ggoSpread.SSSetEdit 		C_LotNo,		"LOTNO",4
			ggoSpread.SSSetEdit 		C_LotSubNo,		"LotSubNo",4
			ggoSpread.SSSetEdit 		C_CostCd,		"Cost Center",4		
			ggoSpread.SSSetFloat		C_SchdQty,		"�������������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat		C_OriginQty,	"������",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
			
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_Status, C_Status, True)
			Call ggoSpread.SSSetColHidden(C_PlantCd, C_OriginQty, True)
			    
			ggoSpread.SSSetSplit2(4)
			
			Call SetSpreadLock("A")
			
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
		If pvSpdNo = "A" Then
			'--------------------------------
			'Grid 1
			'--------------------------------    
			ggoSpread.Source = frm1.vspdData1
			.vspdData1.ReDraw = False
			
			'ggoSpread.SpreadLock -1, -1
			ggoSpread.SpreadLock 2,-1, .vspdData1.maxCols ,-1
			.vspdData1.ReDraw = True
		
		
	   End If
	End With

End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub


'========================== 2.2.6 InitSpreadComboBox()  ========================================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox(ByVal pvSpdNo)
	
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
 Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Then
		' Grid 1(vspdData1) - Production Order
		C_Chk				= 1
		C_ProdtDt			= 2
		C_ProdtOrderNo		= 3 
		C_OprNo				= 4
		C_ItemCd			= 5
		C_ItemNm			= 6
		C_Spec				= 7
		
		C_IssueQty			=8
		C_Remark			=9
		C_GoodQty			=10
		C_BasicUnit			=11
		C_SlCd				=12
		C_SlNm				=13
		C_TrackingNo		=14	
		C_ReqNo				=15
		C_ResvSeq			=16
		C_ResultSeq			=17
		C_DocumentNo		=18
		C_Status			=19
		C_StatusNm			=20
		C_Error				=21
		C_PlantCd			=22
		C_DocumentYear		=23
		C_WcCd				=24
		C_LotNo				=25
		C_LotSubNo			=26
		C_CostCd			=27
		C_SchdQty			=28
		C_OriginQty			=29
		
		
	End If
	
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'=================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Chk						= iCurColumnPos(1)	
			C_ProdtDt					= iCurColumnPos(2)	
			C_ProdtOrderNo				= iCurColumnPos(3)	
			C_OprNo						= iCurColumnPos(4)
			C_ItemCd					= iCurColumnPos(5)
			C_ItemNm					= iCurColumnPos(6)
			C_Spec						= iCurColumnPos(7)
			C_IssueQty					= iCurColumnPos(8)
			C_Remark					= iCurColumnPos(9)
			C_GoodQty					= iCurColumnPos(10)
			C_BasicUnit					= iCurColumnPos(11)
			C_SlCd						= iCurColumnPos(12)
			C_SlNm						= iCurColumnPos(13)
			C_TrackingNo				= iCurColumnPos(14)
			C_ReqNo						= iCurColumnPos(15)
			C_ResvSeq					= iCurColumnPos(16)
			C_ResultSeq					= iCurColumnPos(17)
			C_DocumentNo				= iCurColumnPos(18)
			C_Status					= iCurColumnPos(19)
			C_StatusNm					= iCurColumnPos(20)
			C_Error						= iCurColumnPos(21)
			C_PlantCd					= iCurColumnPos(22)
			C_DocumentYear				= iCurColumnPos(23)
			C_WcCd						= iCurColumnPos(24)
			C_LotNo						= iCurColumnPos(25)
			C_LotSubNo					= iCurColumnPos(26)
			C_CostCd					= iCurColumnPos(27)
			C_SchdQty					= iCurColumnPos(28)
			C_OriginQty					= iCurColumnPos(29)
			
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
'++++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  ------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", fmr1.txtPlantCd.alt,"X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", frm1.txtPlantCd.alt,"X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field��(0)
	arrField(1) = 2 '"ITEM_NM"					' Field��(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", frm1.txtPlantCd.alt,"X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = frm1.txtProdFromDt.Text
	arrParam(4) = frm1.txtProdToDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)

    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With

End Function

'------------------------------------------  SetTrackingNo()  ----------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	
	frm1.txtTrackingNo.Value = arrRet(0)
	
End Function

'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtProdFromDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdFromDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtProdToDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdToDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	


'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdToDt.Focus
    End If
End Sub

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  *********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*******************************************************************************************************

'******************************  3.2.1 Object Tag ó��  ************************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*******************************************************************************************************
Sub vspdData1_Click(ByVal Col , ByVal Row )
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
  	Else
  		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
  	End If
    
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
 		
	Else
 		'------ Developer Coding part (Start)
 	'	If lgOldRow <> Row Then
	'		
			
	'		lgOldRow = Row			
		'End If
	 	'------ Developer Coding part (End)	
 	End If 	
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

'=======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData1_Change(ByVal Col, ByVal Row)

	With frm1.vspdData1		
		Select Case Col
			Case C_IssueQty
				
				ggoSpread.SpreadUnLock C_Remark,Row,C_Remark,Row
				ggoSpread.SSSetRequired C_Remark, Row, Row
			

		End Select 

	End With

End Sub

'==========================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim pvQty, pvRemark
	
    With frm1.vspdData1
		.Row = Row
		.Col = C_Chk
		
		ggoSpread.Source = frm1.vspdData1

		If ButtonDown = 1 Then
			ggoSpread.UpdateRow Row
			ggoSpread.SpreadUnLock C_IssueQty,Row,C_IssueQty,Row
			ggoSpread.SSSetRequired C_IssueQty, Row, Row
		Else
			ggoSpread.SSDeleteFlag Row,Row
			ggoSpread.SpreadLock C_IssueQty,Row,C_IssueQty ,Row
			ggoSpread.SpreadLock C_Remark,Row,C_Remark ,Row
			ggoSpread.SSSetProtected C_IssueQty, Row, Row
			ggoSpread.SSSetProtected C_Remark, Row, Row
			
			Call .GetText(C_OriginQty, Row, pvQty)
			
			Call .SetText(C_IssueQty,		Row, pvQty)
			Call .SetText(C_Remark,		Row, "")
			
		End If			

	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if    
End Sub
'========================================================================================
' Function Name : vspdData1_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

 
'========================================================================================
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData1
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
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

	Dim LngRow
	Dim strInsideFlag
	Dim strMilestoneFlag

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    
	Call ggoSpread.ReOrderingSpreadData()
	
	If gActiveSpdSheet.Id = "A" Then
		Call InitData(1)
		ggoSpread.Source = frm1.vspdData1

	Else
		lgOldRow = 0
		Call vspdData1_Click(frm1.vspdData1.ActiveCol, frm1.vspdData1.ActiveRow)	
    
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
'********************************************************************************************************
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    Dim IntRetCD 
    
    FncQuery = False											'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = displaymsgbox("900013", parent.VB_YES_NO, "x", "x")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function

   '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "3")						'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData

    Call InitVariables

	If DBQuery = False Then Exit Function 

    FncQuery = True												'��: Processing is OK
   
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
    Dim	LngRows
    
    FncSave = False												'��: Processing is NG
    
    Err.Clear                                                   '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = displaymsgbox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
    
    With frm1.vspdData1
		For LngRows = 1 To .MaxRows
			.Row = LngRows
			.Col = C_IssueQty
			If .Value <= 0 Then
				Call DisplayMsgBox("970022", "x", "�Է¼���", "0")
				Call SheetFocus(LngRows, C_IssueQty)
				Exit Function
			End If   
			
			.Col = 0
			If .Text = ggoSpread.DeleteFlag Then
				.Col = C_Remark
				If Trim(.Text) = "" Then
					
					Call DisplayMsgBox("970021", "x", "���", "x")
					Call SheetFocus(LngRows, C_Remark)
					Exit Function
				End If
			
			End If 
		Next 
		
	End With	

    '-----------------------
    'Save function call area
    '-----------------------

    If DbSave = False Then Exit Function						'��: Save db data

    FncSave = True												'��: Processing is OK
    
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
	
    If frm1.vspdData1.MaxRows < 1 Then Exit Function
    ggoSpread.EditUndo 
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

    Dim lDelRows
    Dim pvStatus
    
    If frm1.vspdData1.MaxRows < 1 Then Exit Function
	
	frm1.vspdData1.Col = C_Status
	
	Call frm1.vspdData1.GetText(C_Status, frm1.vspdData1.ActiveRow, pvStatus)
	
	If pvStatus = "D" Or pvStatus = "Y" Then Exit Function
	
	ggoSpread.SpreadUnLock C_Remark,frm1.vspdData1.ActiveRow ,C_Remark, frm1.vspdData1.ActiveRow
	ggoSpread.SSSetRequired C_Remark, frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow
	
	ggoSpread.Source = frm1.vspdData1
    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
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
	
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = displaymsgbox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  ******************************
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

    Err.Clear

    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
	
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)
		strVal = strVal & "&txtProdOrdNo=" & Trim(.hProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.Value)
		strVal = strVal & "&txtProdFromDt=" & Trim(.hProdFromDt.Value)
		strVal = strVal & "&txtProdTODt=" & Trim(.hProdTODt.Value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.Value)
		strVal = strVal & "&txtrdoflag=" & Trim(.hrdoFlag.Value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		
	Else
	
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey		
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4		
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)				
		strVal = strVal & "&txtProdOrdNo=" & Trim(.txtProdOrderNo.Value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)		
		strVal = strVal & "&txtProdFromDt=" & Trim(.txtProdFromDt.Text)		
		strVal = strVal & "&txtProdTODt=" & Trim(.txtProdTODt.Text)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.Value)
		
		If frm1.rdoCompleteFlg1.checked = True Then		
			strVal = strVal & "&txtrdoflag=" & "N"
		ElseIf  frm1.rdoCompleteFlg2.checked = True  Then
			strVal = strVal & "&txtrdoflag=" & "Y"
		ElseIf  frm1.rdoCompleteFlg4.checked = True  Then	
			strVal = strVal & "&txtrdoflag=" & "D"
		Else
			strVal = strVal & "&txtrdoflag=" & "A"
		End If
		
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)

    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Dim lRow
	Dim strStatus
	Dim strMilestoneFlag
	Dim DblInvQty
	
	Call InitData(LngMaxRow)

	frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1

	lgOldRow = 1

    With frm1.vspdData1

		.Redraw = False
		
		ggoSpread.Source = frm1.vspdData1
		
		If .MaxRows >0 and frm1.hrdoFlag.value<>"Y" Then
			frm1.btnRun.disabled=False
		Else
			frm1.btnRun.disabled=True
		End If		
	
		For lRow = LngMaxRow To .MaxRows

			ggoSpread.Source = frm1.vspdData1
			
			.Row = lRow
			.Col = C_Status
			strStatus = Trim(.text)

			
			.Col = C_GoodQty
			DblInvQty = uniCDbl(.text)
			
			If UCase(strStatus) = "Y" OR UCase(strStatus) = "D" Then
				ggoSpread.SpreadLock C_Chk,lRow,C_Chk ,lRow
			Else
				If DblInvQty > uniCDbl(0) Then
					ggoSpread.SpreadUnLock C_Chk,lRow,C_Chk,lRow
				Else 
					ggoSpread.SpreadLock C_Chk,lRow,C_Chk ,lRow	
				End If	
			End If	
		Next
		
		.Redraw = True
    
    End With

	Call SetToolBar("11001011000111")										'��: ��ư ���� ���� 

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement		

	End If
	

	lgIntFlgMode = parent.OPMD_UMODE

End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryNotOk()

	Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 	
	frm1.btnRun.disabled=True	
	
End Function
'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim strVal, strDel  
    Dim IntRows
    
    Dim lGrpCnt
	
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen						'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

	Dim iTmpDBuffer							'������ ���� [����] 
	Dim iTmpDBufferCount					'������ ���� Position
	Dim iTmpDBufferMaxCount					'������ ���� Chunk Size
	
	lGrpCnt = 1

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0
	
    DbSave = False                                                          '��: Processing is NG
    
    Call LayerShowHide(1)
    
	With frm1.vspdData1

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0

			Select Case .Text
		    
			    Case ggoSpread.UpdateFlag
			    
					
					strVal = ""
					.Col = C_ProdtOrderNo	
					strVal = strVal & UCase(Trim(.Text)) & iColSep		' 0
					.Col = C_OprNo	
					strVal = strVal & UCase(Trim(.Text)) & iColSep		' 1
					.Col = C_ResvSeq		
					strVal = strVal & Trim(.Text) & iColSep					' 2 
					.Col = C_ResultSeq	
					strVal = strVal & Trim(.Text) & iColSep					' 3   
					.Col = C_PlantCd		
					strVal = strVal & Trim(.Text) & iColSep					' 4
					.Col = C_ItemCd	
						
					strVal = strVal & UCase(Trim(.Text)) & iColSep		' 5
					strVal = strVal & "" & iColSep								'resv_type 6
					.Col = C_TrackingNo		
					strVal = strVal & Trim(.Text) & iColSep					'7
					
					.Col = C_IssueQty		
					strVal = strVal & UNIConvNum(.Text,0) & iColSep	' Issue Qty	8
					
					If uniCDbl(.Text) <= 0 Then
						
						Exit Function
					End If
					
					.Col = C_SchdQty		
					strVal = strVal & UNIConvNum(.Text,0) & iColSep	' SCHD Qty	9
					.Col = C_BasicUnit		
					strVal = strVal & UCase(Trim(.Text)) & iColSep		' 10
					.Col = C_SlCd		
					strVal = strVal & UCase(Trim(.Text)) & iColSep		' 11
					strVal = strVal & "A" & iColSep							'issue mthd 12
					.Col = C_WcCd		
					strVal = strVal & UCase(Trim(.Text)) & iColSep		' wc code   13 
					.Col = C_ReqNo
					strVal = strVal & Trim(.Text) & iColSep 				'14		
					.Col = C_ProdtDt	
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep	' REPORT Date		15						
					.Col = C_LotNo
					strVal = strVal & Trim(.Text) & iColSep					' 16
					.Col = C_LotSubNo
					strVal = strVal & Trim(.Text) & iColSep					' 17
					.Col = C_CostCd
					strVal = strVal & Trim(.Text) & iColSep					' 18  	
					
					.Col = C_DocumentNo
					strVal = strVal & Trim(.Text) & iColSep					' 19
					.Col = C_DocumentYear
					strVal = strVal & Trim(.Text) & iColSep    				' 20	
					.Col = C_Status
					strVal = strVal & Trim(.Text) & iColSep					' 21
					.Col = C_Error
					strVal = strVal & "" & iColSep					' ERROR DESC 22
					.Col = C_OriginQty
					strVal = strVal & UNIConvNum(.Text,0) & iColSep					' ERROR DESC 22 					
					.Col = C_Remark
					strVal = strVal & Trim(.Text) & iColSep					' ERROR DESC 22 
					
					'strVal = strVal & iColSep
					strVal = strVal & IntRows & iRowSep
			
					lGrpCnt = lGrpCnt + 1
					 
				Case ggoSpread.DeleteFlag	 
					
					strDel = ""
					.Col = C_ProdtOrderNo	
					strDel = strDel & UCase(Trim(.Text)) & iColSep		' 0
					.Col = C_OprNo	
					strDel = strDel & UCase(Trim(.Text)) & iColSep		' 1
					.Col = C_ResvSeq		
					strDel = strDel & Trim(.Text) & iColSep					' 2 
					.Col = C_ResultSeq	
					strDel = strDel & Trim(.Text) & iColSep					' 3   
					.Col = C_PlantCd		
					strDel = strDel & Trim(.Text) & iColSep					' 4
					.Col = C_ItemCd	
						
					strDel = strDel & UCase(Trim(.Text)) & iColSep		' 5
					strDel = strDel & "" & iColSep								'resv_type 6
					.Col = C_TrackingNo		
					strDel = strDel & Trim(.Text) & iColSep					'7

					.Col = C_IssueQty		
					strDel = strDel & UNIConvNum(.Text,0) & iColSep	' Issue Qty	8
					.Col = C_SchdQty		
					strDel = strDel & UNIConvNum(.Text,0) & iColSep	' SCHD Qty	9
					.Col = C_BasicUnit		
					strDel = strDel & UCase(Trim(.Text)) & iColSep		' 10
					.Col = C_SlCd		
					strDel = strDel & UCase(Trim(.Text)) & iColSep		' 11
					strDel = strDel & "A" & iColSep							'issue mthd 12
					.Col = C_WcCd		
					strDel = strDel & UCase(Trim(.Text)) & iColSep		' wc code   13 
					.Col = C_ReqNo
					strDel = strDel & Trim(.Text) & iColSep 				'14		
					.Col = C_ProdtDt	
					strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep	' REPORT Date		15						
					.Col = C_LotNo
					strDel = strDel & Trim(.Text) & iColSep					' 16
					.Col = C_LotSubNo
					strDel = strDel & Trim(.Text) & iColSep					' 17
					.Col = C_CostCd
					strDel = strDel & Trim(.Text) & iColSep					' 18  	
					
					.Col = C_DocumentNo
					strDel = strDel & Trim(.Text) & iColSep					' 19
					.Col = C_DocumentYear
					strDel = strDel & Trim(.Text) & iColSep    				' 20	
					.Col = C_Status
					strDel = strDel & "D" & iColSep					' 21
					.Col = C_Error
					strDel = strDel & "" & iColSep					' ERROR DESC 22
					.Col = C_OriginQty
					strDel = strDel & UNIConvNum(.Text,0) & iColSep					' ERROR DESC 22 					
					.Col = C_Remark
					strDel = strDel & Trim(.Text) & iColSep					' ERROR DESC 22 
					
					'strDel = strDel & iColSep
					strDel = strDel & IntRows & iRowSep
			
					lGrpCnt = lGrpCnt + 1
					
			End Select
			
			
			.Col = 0
			Select Case .Text
			    Case ggoSpread.UpdateFlag
			    
			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			         
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
			
	    Next
	    
	End With
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  
	   
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)

    DbSave = True
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()
   
    lgIntPrevKey = 0
    lgLngCurRows = 0

	ggoSpread.source = frm1.vspddata1
    frm1.vspdData1.MaxRows = 0
	lgIntFlgMode = parent.OPMD_CMODE
	
	Call RemovedivTextArea
	Call DbQuery
	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
    On Error Resume Next
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData1.focus
	frm1.vspdData1.Row = lRow
	frm1.vspdData1.Col = lCol
	frm1.vspdData1.Action = 0
	frm1.vspdData1.SelStart = 0
	frm1.vspdData1.SelLength = len(frm1.vspdData1.Text)

	'Call RestoreToolBar()

End Function
'========================== 2.2.6 InitComboBox()  ========================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

End Sub

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)        
	    
    Call ggoOper.LockField(Document, "Q")                                   '��: Lock  Suitable  Field
    Call InitSpreadSheet("A")                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal

    Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If	
End Sub
'========================================================================================
' Function Name : Checkall()
'========================================================================================
Function Checkall()
	
 Dim IRowCount 
 Dim IClnCount
 Dim iDblQty
 
 
 ggoSpread.Source = frm1.vspdData1
 With frm1.vspdData1 
 
	.Redraw = False
    
	IF lgCheckall = 0 Then		'select All

		For IRowCount = 1 to .MaxRows 	     	 
			.Row = IRowCount 
			.Col = C_Status
						
			If  Ucase(Trim(.Text)) <>"Y" AND Ucase(Trim(.Text)) <>"D"  Then 
				.Col = C_GoodQty
				If uniCDbl(.Text) > 0 Then
					.Col = C_Chk	 
					.text = 1   
				End If	  
			End If
		Next    

		
		frm1.btnRun.value = "��ü�������"
		
		lgCheckall = 1
		lgBlnFlgChgValue = True
	 Else							'deselect All
	  
		 For IRowCount = 1 to .MaxRows 
			.Row = IRowCount 
			.Col = C_Status
			If  Ucase(Trim(.Text)) <>"Y" AND Ucase(Trim(.Text)) <>"D"  Then 
				.Col = C_GoodQty
		    	If uniCDbl(.Text) > 0 Then
		    		.Col = C_Chk	 
					.text = 0 
				End If	  
			End If
		NEXT
		
		frm1.btnRun.value = "��ü����"
		
		lgCheckall = 0
		lgBlnFlgChgValue = False
	End If
	
	.Redraw = True

 End With
 
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS="TD6">
										<script language=JavaScript>
										ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtProdFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');
										</script>
										&nbsp;~&nbsp;
										<script language=JavaScript>
										ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtProdTODt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');
										</script>
									</TD>																						
								</TR>
								<TR>
								<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>	
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>								
								</TR>
								<TR>									
									<TD CLASS=TD5 NOWRAP>ó������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoCompleteFlg" ID="rdoCompleteFlg1" CLASS="RADIO" tag="11" Value="N" CHECKED><LABEL FOR="rdoCompleteFlg1">��ó��</LABEL>
									     				 <INPUT TYPE="RADIO" NAME="rdoCompleteFlg" ID="rdoCompleteFlg2" CLASS="RADIO" tag="11" Value="Y" ><LABEL FOR="rdoCompleteFlg2">ó��</LABEL>
									     				 <INPUT TYPE="RADIO" NAME="rdoCompleteFlg" ID="rdoCompleteFlg3" CLASS="RADIO" tag="11" Value="A" ><LABEL FOR="rdoCompleteFlg3">��ü</LABEL>
									     				 <INPUT TYPE="RADIO" NAME="rdoCompleteFlg" ID="rdoCompleteFlg4" CLASS="RADIO" tag="11" Value="D" ><LABEL FOR="rdoCompleteFlg4">����</LABEL></TD>
  									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
			
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="100%">
								<TD WIDTH="100%" colspan=4>
								<script language=JavaScript>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 ID = "A" width="100%" tag="3" TITLE="SPREAD" id=OBJECT1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');
								</script>
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
		<TD  HEIGHT=3></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:Checkall()" disabled>��ü ����</BUTTON></TD>		
					<TD WIDTH=10>&nbsp;</TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdTODt" tag="24">
<INPUT TYPE=HIDDEN NAME="hrdoFlag" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
