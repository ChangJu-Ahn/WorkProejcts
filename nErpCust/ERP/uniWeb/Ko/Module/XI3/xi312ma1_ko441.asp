
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name		: INTERFACE 
'*  2. Function Name		: 
'*  3. Program ID		: xi312ma1_KO441
'*  4. Program Name		: ����������� 
'*  5. Program Desc		:
'*  6. Comproxy List		: +B19029LookupNumericFormat
'*  7. Modified date(First)	: 2006/04/24
'*  8. Modified date(Last) 	: 
'*  9. Modifier (First) 	: HJO
'* 10. Modifier (Last)		: 
'* 11. Comment			:
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
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

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim LocSvrDate
Dim StartDate
Dim EndDate

	iDBSYSDate = "<%=GetSvrDate%>"											'��: DB�� ���� ��¥�� �޾ƿͼ� ���۳�¥�� ����Ѵ�.
	LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	StartDate = UNIDateAdd("D",-1,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
	EndDate = UNIDateAdd("D", 0,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

<!-- #Include file="../../inc/lgvariables.inc" -->
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID = "xi312mb1_KO441.asp"											'��: �����Ͻ� ����(Qeury) ASP�� 
'Post Production Results
Const BIZ_PGM_SAVE_ID	= "xi312mb2_KO441.asp"		
Const BIZ_PGM_SAVE_ID2	= "xi312mb3_KO441.asp"
Const BIZ_PGM_SAVE_ID3	= "xi312mb4_KO441.asp"
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Dim C_ProdtOrderNo
Dim C_ResultIfSeq
Dim C_ReportDt
Dim C_ItemCd		
Dim C_ItemNm
Dim C_LotNo
Dim C_report_type
Dim C_ProdQtyInOrderUnit
Dim C_loss_qty
Dim C_bonus_qty
Dim C_JobStatus
Dim C_JobTime
Dim C_JobLine
Dim C_CreateType
Dim C_TagNo
Dim C_ResultSeq	
Dim C_SendDt
Dim C_ReceiveDt	
Dim C_ApplyFlag	
Dim C_ErrDesc

Dim C_PopInfKey
Dim C_PlantCd
'2008-06-23 2:07���� :: hanc
Dim C_WcCd
Dim C_ToOper

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgStrPrevKey5
Dim lgStrPrevKey6

Dim lgOldRow
Dim lgLngCnt
Dim lgAfterQryFlg

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>

End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitComboBox()
			
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) 

    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
         
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
			
    Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables

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

'20080122::hanc	frm1.btnReg.disabled = true
End Sub

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop          
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

    lgIntFlgMode = parent.OPMD_CMODE	'Indicates that current mode is Create mode
    lgIntGrpCount = 0			'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""			'initializes Previous Key
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgStrPrevKey5 = ""
    lgStrPrevKey6 = ""
    lgLngCurRows = 0		'initializes Deleted Rows Count
    
    lgLngCnt = 0
    lgOldRow = 0
    lgAfterQryFlg = False
    lgSortKey    = 1
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
	frm1.txtSendStartDt.Text = EndDate
	frm1.txtSendEndDt.Text = EndDate
	frm1.txtPlanStartDt.Text = EndDate
	frm1.txtPlanEndDt.Text = EndDate
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	
	With frm1.vspdData 
		 
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20060418",,parent.gAllowDragDropSpread    
		
		.ReDraw = false
	
		.MaxCols = C_ToOper + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ    
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit		C_ProdtOrderNo,			"����������ȣ", 15, , ,18
		ggoSpread.SSSetEdit		C_TagNo,				"TAG No",		15
		ggoSpread.SSSetEdit		C_ResultIfSeq,			"�������ۼ���",	10, 2 
		ggoSpread.SSSetDate 	C_ReportDt, 			"��������",		10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ItemCd,				"ǰ��",			15, , ,18
		ggoSpread.SSSetEdit		C_ItemNm,				"ǰ���",		25, , ,40
		ggoSpread.SSSetEdit		C_LotNo,				"LOT��ȣ",		20, , ,25
		ggoSpread.SSSetEdit		C_report_type,				"REPORT TYPE",		11, , ,25
		ggoSpread.SSSetFloat	C_ProdQtyInOrderUnit,	"��������",		11,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_loss_qty,	"LOSS ����",		11,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_bonus_qty,	"BONUS����",		11,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_JobStatus, 			"�۾�����",		 8,2
		ggoSpread.SSSetEdit 	C_JobTime, 				"�����ð�",		 8,2
		ggoSpread.SSSetEdit		C_JobLine,				"����",			 8,2
		ggoSpread.SSSetEdit		C_CreateType,			"��������",		 8,2			
		ggoSpread.SSSetEdit		C_ResultSeq,			"ERP��������",	10,2
		ggoSpread.SSSetEdit		C_SendDt ,				"MES�۽��Ͻ�" , 20
		ggoSpread.SSSetEdit		C_ReceiveDt ,			"ERP�����Ͻ�" , 20
		ggoSpread.SSSetEdit		C_ApplyFlag,			"ERP�ݿ�����",	10,2		
		ggoSpread.SSSetEdit 	C_ErrDesc, 				"��������",		30
		ggoSpread.SSSetEdit		C_PopInfKey,			"PopKey",		6
		ggoSpread.SSSetEdit		C_PlantCd,				"����",			6
		ggoSpread.SSSetEdit		C_WcCd, 				"�۾���",		13
		ggoSpread.SSSetEdit		C_ToOper,				"To OPER",			13


				
		Call ggoSpread.SSSetColHidden(C_TagNo, C_ResultIfSeq, True)
		Call ggoSpread.SSSetColHidden(C_JobTime, C_JobLine, True)
		Call ggoSpread.SSSetColHidden(C_ResultSeq, C_SendDt, True)
		Call ggoSpread.SSSetColHidden(C_PopInfKey, C_PlantCd, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)         
'		Call ggoSpread.SSSetColHidden(C_JobLine, .MaxCols, True)          '20080225::hanc
	
		ggoSpread.SSSetSplit2(3)							'frozen ����߰� 
	
		Call SetSpreadLock 
    
		.ReDraw = true
    
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
	
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
End Sub


'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitSpreadComboBox()
	
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex
	
	With frm1.vspdData

	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	
	C_ProdtOrderNo			= 1
	C_TagNo					= 2
	C_ResultIfSeq		    = 3
	C_ReportDt				= 4
	C_ItemCd		        = 5
	C_ItemNm				= 6	
	C_LotNo					= 7
	C_report_type   		= 8
	C_ProdQtyInOrderUnit	= 9
	C_loss_qty	            = 10
	C_bonus_qty	            = 11
	C_JobStatus				= 12
	C_JobTime				= 13
	C_JobLine				= 14
	C_CreateType			= 15
	C_ResultSeq				= 16
	C_SendDt				= 17
	C_ReceiveDt				= 18
	C_ApplyFlag				= 19
	C_ErrDesc				= 20
	C_PopInfKey				= 21
	C_PlantCd				= 22
	C_WcCd				    = 23
	C_ToOper				= 24
	
End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 				C_ProdtOrderNo			= iCurColumnPos(1)
				C_TagNo					= iCurColumnPos(2)
				C_ResultIfSeq			= iCurColumnPos(3)
				C_ReportDt				= iCurColumnPos(4)
				C_ItemCd				= iCurColumnPos(5)
				C_ItemNm				= iCurColumnPos(6)
				C_LotNo					= iCurColumnPos(7)
				C_report_type	        = iCurColumnPos(8)
				C_ProdQtyInOrderUnit	= iCurColumnPos(9) 
				C_loss_qty	            = iCurColumnPos(10)
				C_bonus_qty	            = iCurColumnPos(11)
				C_JobStatus				= iCurColumnPos(12)
				C_JobTime				= iCurColumnPos(13)
				C_JobLine				= iCurColumnPos(14)
				C_CreateType			= iCurColumnPos(15)
				C_ResultSeq				= iCurColumnPos(16)
				C_SendDt				= iCurColumnPos(17)
				C_ReceiveDt				= iCurColumnPos(18)
				C_ApplyFlag				= iCurColumnPos(19)
				C_ErrDesc				= iCurColumnPos(20)
				C_PopInfKey				= iCurColumnPos(21)
				C_PlantCd				= iCurColumnPos(22)
				C_WcCd				    = iCurColumnPos(23)
				C_ToOper				= iCurColumnPos(24)

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

'------------------------------------------  OpenPlant()  -------------------------------------------------
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
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNo()
    
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
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
	arrParam(3) = ""
	arrParam(4) = ""
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

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode)
    
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
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
	arrParam(1) = strCode			' Item Code
	arrParam(2) = ""				' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field��(0)
	arrField(1) = 2 '"ITEM_NM"			' Field��(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function
'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)		
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
'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtSendStartDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtSendStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtSendEndDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtSendEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================

Sub txtSendStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSendStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtSendStartDt.Focus
    End If
End Sub

Sub txtSendEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSendEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtSendEndDt.Focus
    End If
End Sub
'------------------------------------------  txtPlanStartDt_KeyDown ----------------------------------------
'	Name : txtPlanStartDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlanStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtPlanEndDt_KeyDown ------------------------------------------
'	Name : txtPlanEndDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlanEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================

Sub txtPlanStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanStartDt.Focus
    End If
End Sub

Sub txtPlanEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanEndDt.Focus
    End If
End Sub
'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	
 	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
 	
 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData
    
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
 		Exit Sub
 	End If
 	
' 	If Row <= 0 Then
'       ggoSpread.Source = frm1.vspdData
'       Exit Sub
'    End If

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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    'Exit Sub
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "" And lgStrPrevKey4 <> "" And _
		   lgStrPrevKey5 <> ""  And lgStrPrevKey6 <> ""  Then				'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
    
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
	
	If ValidDateCheck(frm1.txtSendStartDt, frm1.txtSendEndDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtPlanStartDt, frm1.txtPlanEndDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If																'��: Query db data

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
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow
    
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
	With frm1
    End With
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

Function FncPrint()                                                  '��: Protect system from crashing
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

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    
    ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)
    
End Sub 


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
	Dim rdoFlag
	Dim rdoFlagF
    
    DbQuery = False
  
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing
        
    With frm1
    
    If .rdoFlg1.checked then
		rdoFlag="A"
	ElseIf .rdoFlg2.checked Then 
		rdoFlag ="Y"
	Else
		rdoFlag ="N"
	End If
    
    If .rdoFlgF1.checked then
		rdoFlagF="A"
	ElseIf .rdoFlgF2.checked Then 
		rdoFlagF ="Y"
	Else
		rdoFlagF ="N"
	End If
    

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtPlantCd="       & Trim(.hPlantCd.Value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtToOper="       & Trim(.txtToOper.Value)				'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&txtWcCd="       & Trim(.txtWcCd.Value)				'��: ��ȸ ���� ����Ÿ	

		strVal = strVal & "&txtSendStartDt="   & Trim(.hSendStartDt.Value)
		strVal = strVal & "&txtSendEndDt="     & Trim(.hSendEndDt.Value)
		strVal = strVal & "&txtPlanStartDt="   & Trim(.hPlanStartDt.value)
		strVal = strVal & "&txtPlanEndDt="     & Trim(.hPlanEndDt.value)
		strVal = strVal & "&txtProdtOrderNo="  & Trim(.hProdOrderNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd="        & Trim(.hItemCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&rdoFlag="          & Trim(.hRdoFlag.Value)				'��: ��ȸ ���� ����Ÿ 
		
	Else
		strVal =  BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&txtPlantCd="       & Trim(.txtPlantCd.Value)				'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&txtToOper="       & Trim(.txtToOper.Value)				'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&txtWcCd="       & Trim(.txtWcCd.Value)				'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&txtSendStartDt="   & Trim(.txtSendStartDt.Text)
		strVal = strVal & "&txtSendEndDt="     & Trim(.txtSendEndDt.Text)
		strVal = strVal & "&txtPlanStartDt="   & Trim(.txtPlanStartDt.Text)
		strVal = strVal & "&txtPlanEndDt="     & Trim(.txtPlanEndDt.Text)
		strVal = strVal & "&txtProdtOrderNo="  & Trim(.txtProdOrderNo.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd="        & Trim(.txtItemCd.Value)		'��: ��ȸ ���� ����Ÿ 		
		strVal = strVal & "&rdoFlag="          & Trim(rdoFlag)				'��: ��ȸ ���� ����Ÿ 
	End If
	strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
	strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
	strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
	strVal = strVal & "&lgStrPrevKey5=" & lgStrPrevKey5
	strVal = strVal & "&lgStrPrevKey6=" & lgStrPrevKey6
	strVal = strVal & "&rdoFlagF="          & Trim(rdoFlagF)				'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)											'��: ��ȸ ������ ������� 
	
	Call SetToolBar("11000000000111")										'��: ��ư ���� ���� 
    '-----------------------
    'Reset variables area
    '-----------------------
    
    
    Call InitData(LngMaxRow)
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    
    
	
End Function



'========================================================================================
' Function Name : RecMes
' Function Desc : This function is data query and display
'========================================================================================
Function RecMes() 


    Dim IntRows 
    Dim strVal  
	
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

    RecMes = False
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------  
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)

    RecMes = True
    
End Function
'========================================================================================
' Function Name : RegProd
' Function Desc : This function is update 
'========================================================================================
Function RegProd() 

    Dim IntRows 
    Dim strVal  
	
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

    RegProd = False
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
	End With

  	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID2)

    RegProd = True
    
End Function

Function RegProd1() 

    Dim IntRows 
    Dim strVal  
	
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

    RegProd1 = False
    
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
	End With

  	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID3)

    RegProd1 = True
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()
   
    lgIntPrevKey = 0
    lgLngCurRows = 0

	ggoSpread.source = frm1.vspddata
    frm1.vspddata.MaxRows = 0
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

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>MES�۽űⰣ</TD>
								    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtSendStartDt 	CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtSendEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
								</TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>									
									<TD CLASS=TD5 NOWRAP>��������Ⱓ</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtPlanStartDt	CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="11X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtPlanEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="11X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
								</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>TO_OPER</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtToOper" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="To Oper">&nbsp;</TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="WcCd">&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>����ó������</TD>
									<TD CLASS=TD6 NOWRAP>
									    <INPUT TYPE="RADIO" NAME="rdoFlgF" ID="rdoFlgF1" CLASS="RADIO" tag="11" Value="A" CHECKED><LABEL FOR="rdoFlgF1">��ü</LABEL>
									    <INPUT TYPE="RADIO" NAME="rdoFlgF" ID="rdoFlgF2" CLASS="RADIO" tag="11" Value="Y" ><LABEL FOR="rdoFlgF2">����ó��</LABEL>
				     				    <INPUT TYPE="RADIO" NAME="rdoFlgF" ID="rdoFlgF3" CLASS="RADIO" tag="11" Value="N" ><LABEL FOR="rdoFlgF3">ó���߿���</LABEL></TD>
									</TD>								
								</TR>								
								<TR style="display:none">	
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP>ERP�ݿ�����</TD>
									<TD CLASS=TD6 NOWRAP>
									    	<INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg1" CLASS="RADIO" tag="11" Value="A" CHECKED><LABEL FOR="rdoFlg1">��ü</LABEL>
				     				    <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg2" CLASS="RADIO" tag="11" Value="Y" ><LABEL FOR="rdoFlg2">�ݿ�</LABEL>
				     				    <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg3" CLASS="RADIO" tag="11" Value="N" ><LABEL FOR="rdoFlg3">�̹ݿ�</LABEL>
									</TD>								
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
						<TABLE <%=LR_SPACE_TYPE_60%>>								
							<TR HEIGHT="100%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPSPD%> HEIGHT=100% NAME=vspdData WIDTH=100% tag="2" TITLE="SPREAD" id=OBJECT2><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');
									</SCRIPT>
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
				<TR style="display:none">
					<TD>&nbsp;</TD>
					<TD><!--<BUTTON NAME="btnRec" CLASS="CLSMBTN" ONCLICK="vbscript:RecMes()" >MES��������</BUTTON> -->
					&nbsp;<BUTTON NAME="btnReg" CLASS="CLSMBTN" ONCLICK="vbscript:RegProd()" >����������</BUTTON>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;<BUTTON NAME="btnReg" CLASS="CLSMBTN" ONCLICK="vbscript:RegProd1()" >ó���߿���DATA���</BUTTON></TD>				
					<TD>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hSendStartDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hSendEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlanStartDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlanEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hRdoFlag" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
