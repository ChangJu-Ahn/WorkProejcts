<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name				: Production
'*  2. Function Name			: 
'*  3. Program ID				: p4212ma1
'*  4. Program Name         	: Release By Order
'*  5. Program Desc        		:
'*  6. Comproxy List        	: +
'*  7. Modified date(First) 	: 2000/03/30
'*  8. Modified date(Last)  	: 2002/07/18
'*  9. Modifier (First)     	: Kim, GyoungDon
'* 10. Modifier (Last)      	: Chen, Jae Hyun
'* 11. Comment              	:
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
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

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID	= "p44B2qb1.asp"							'��: �����Ͻ� ����(Qeury) ASP�� 
Const BIZ_PGM_QRY2_ID	= "p44B2qb2.asp"							'��: �����Ͻ� ����(Save) ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

' vspdData1

Dim C_pop_inf_key
Dim C_erp_apply_flag
Dim C_updt_user_id
Dim C_updt_dt
Dim C_err_desc


Dim C_input_type2
Dim C_prodt_order_no2
Dim C_item_cd2
Dim C_item_nm2
Dim C_spec2
Dim C_opr_no2
Dim C_result_seq2
Dim C_report_type2
Dim C_pop_qty2
Dim C_pop_unit2
Dim C_prodt_order_qty2
Dim C_prod_qty2
Dim C_report_dt2
Dim C_shift_cd2
Dim C_description2
Dim C_reason_cd2
Dim C_reason_nm2
Dim C_tracking_no2
Dim C_lot_no2
Dim C_lot_sub_no2
Dim C_rcpt_item_document_no2
Dim C_remark2
Dim C_subcontract_prc2
Dim C_subcontract_amt2
Dim C_cur_cd2
Dim C_item_group_cd2
Dim C_item_group_nm2


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgOldRow1

Dim lgSortKey1
Dim lgSortKey2

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop          

Dim lgButtonSelection

Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)													
StartDate = UNIDateAdd("D",-10,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
EndDate = UNIDateAdd("D", 20,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
    
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet("*")                                               '��: Setup the Spread sheet

    frm1.txtFromDt.Text = StartDate
    frm1.txtToDt.Text   = EndDate

	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables

    Call InitComboBox
    Call InitSpreadComboBox
    Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtInfNo.focus 
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If
	
End Sub

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
    lgStrPrevKey = ""			'initializes Previous Key
    lgLngCurRows = 0		'initializes Deleted Rows Count
    lgOldRow1 = 0
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


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	
	Call InitSpreadPosVariables(pvSpdNo) 
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then

		With frm1.vspdData1
	
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20041206", , Parent.gAllowDragDropSpread
			.ReDraw = false
	
			.MaxCols = C_err_desc + 1
			.MaxRows = 0
	
			Call GetSpreadColumnPos("A")
	
			ggoSpread.SSSetEdit		C_pop_inf_key,		"�Է¹�ȣ", 18
			ggoSpread.SSSetEdit		C_erp_apply_flag,	"�ݿ�����", 8
			ggoSpread.SSSetEdit		C_updt_user_id,		"�ݿ���", 10
			ggoSpread.SSSetEdit		C_updt_dt,			"�ݿ�����",18
			ggoSpread.SSSetEdit		C_err_desc,			"��������", 100
			
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
	
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("A")
			
			.ReDraw = true

		End With
	
	End If
	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then

		With frm1.vspdData2
	
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20041206", , Parent.gAllowDragDropSpread
			.ReDraw = false
	
			.MaxCols = C_item_group_nm2 + 1
			.MaxRows = 0
	
			Call GetSpreadColumnPos("B")
	
			ggoSpread.SSSetEdit		C_input_type2,		"�Է�����", 8
			ggoSpread.SSSetEdit		C_prodt_order_no2,	"����������ȣ",18
			ggoSpread.SSSetEdit		C_item_cd2,			"ǰ��",18
			ggoSpread.SSSetEdit		C_item_nm2,			"ǰ���",25
			ggoSpread.SSSetEdit		C_spec2,			"���", 25
			ggoSpread.SSSetEdit		C_opr_no2,			"����", 8
			ggoSpread.SSSetEdit		C_result_seq2,		"��������", 8
			ggoSpread.SSSetEdit		C_report_type2,		"��/��",6
			ggoSpread.SSSetFloat	C_pop_qty2,			"�ݿ�����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_pop_unit2,			"��������", 8
			ggoSpread.SSSetFloat	C_prodt_order_qty2,	"��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_prod_qty2,			"��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetDate		C_report_dt2,		"������", 10, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_shift_cd2,			"Shift", 8
			ggoSpread.SSSetEdit		C_description2,		"Shift��", 12
			ggoSpread.SSSetEdit		C_reason_cd2,		"�ҷ�����", 12
			ggoSpread.SSSetEdit		C_reason_nm2,		"�ҷ�����", 12
			ggoSpread.SSSetEdit		C_Tracking_No2,		"Tracking No.", 25
			ggoSpread.SSSetEdit		C_lot_no2,			"Lot No.", 18
			ggoSpread.SSSetEdit		C_lot_sub_no2,		"����", 8
			ggoSpread.SSSetEdit		C_rcpt_item_document_no2,"�԰��ȣ", 20
			ggoSpread.SSSetEdit		C_remark2,			"���", 12
			ggoSpread.SSSetFloat	C_subcontract_prc2,	"���ִܰ�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_subcontract_amt2,	"���ֱݾ�", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_cur_cd2,			"������ȭ", 8
			ggoSpread.SSSetEdit 	C_item_group_cd2,	"ǰ��׷�",	15
			ggoSpread.SSSetEdit		C_item_group_nm2,	"ǰ��׷��", 30
			
			'Call ggoSpread.MakePairsColumn(,)
			Call ggoSpread.SSSetColHidden( C_reason_cd2, C_reason_cd2, True)
			Call ggoSpread.SSSetColHidden( C_subcontract_prc2, C_cur_cd2, True)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
	
			ggoSpread.SSSetSplit2(3	)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true

		End With
	
	End If
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    
    With frm1
		If pvSpdNo = "A" Then
			ggoSpread.Source = .vspdData1
			ggoSpread.SpreadLock -1, -1	' Set Lock Property : Spread 1
		End If	
		
		If pvSpdNo = "B" Then
			ggoSpread.Source = .vspdData2
			ggoSpread.SpreadLock -1, -1	' Set Lock Property : Spread 1
		End If	
	End With
	
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
End Sub


'==========================================  2.2.6 InitSpreadComboBox()  =================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()

End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)

End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' vspdData1
		C_pop_inf_key				= 1
		C_erp_apply_flag			= 2
		C_updt_user_id				= 3
		C_updt_dt					= 4
		C_err_desc					= 5
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' vspdData2
		C_input_type2				= 1
		C_prodt_order_no2			= 2
		C_item_cd2					= 3
		C_item_nm2					= 4
		C_spec2						= 5
		C_opr_no2					= 6
		C_result_seq2				= 7
		C_report_type2				= 8
		C_pop_qty2					= 9
		C_pop_unit2					= 10
		C_prodt_order_qty2			= 11
		C_prod_qty2					= 12
		C_report_dt2				= 13
		C_shift_cd2					= 14
		C_description2				= 15
		C_reason_cd2				= 16
		C_reason_nm2				= 17
		C_tracking_no2				= 18
		C_lot_no2					= 19
		C_lot_sub_no2				= 20
		C_rcpt_item_document_no2	= 21
		C_remark2					= 22
		C_subcontract_prc2			= 23
		C_subcontract_amt2			= 24
		C_cur_cd2					= 25
		C_item_group_cd2			= 26
		C_item_group_nm2			= 27
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
			
			C_pop_inf_key				= iCurColumnPos(1)
			C_erp_apply_flag			= iCurColumnPos(2)
			C_updt_user_id				= iCurColumnPos(3)
			C_updt_dt					= iCurColumnPos(4)
			C_err_desc					= iCurColumnPos(5)
		Case "B"
 			ggoSpread.Source = frm1.vspdData2
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_input_type2				= iCurColumnPos(1)
			C_prodt_order_no2			= iCurColumnPos(2)
			C_item_cd2					= iCurColumnPos(3)
			C_item_nm2					= iCurColumnPos(4)
			C_spec2						= iCurColumnPos(5)
			C_opr_no2					= iCurColumnPos(6)
			C_result_seq2				= iCurColumnPos(7)
			C_report_type2				= iCurColumnPos(8)
			C_pop_qty2					= iCurColumnPos(9)
			C_pop_unit2					= iCurColumnPos(10)
			C_prodt_order_qty2			= iCurColumnPos(11)
			C_prod_qty2					= iCurColumnPos(12)
			C_report_dt2				= iCurColumnPos(13)
			C_shift_cd2					= iCurColumnPos(14)
			C_description2				= iCurColumnPos(15)
			C_reason_cd2				= iCurColumnPos(16)
			C_reason_nm2				= iCurColumnPos(17)
			C_tracking_no2				= iCurColumnPos(18)
			C_lot_no2					= iCurColumnPos(19)
			C_lot_sub_no2				= iCurColumnPos(20)
			C_rcpt_item_document_no2	= iCurColumnPos(21)
			C_remark2					= iCurColumnPos(22)
			C_subcontract_prc2			= iCurColumnPos(23)
			C_subcontract_amt2			= iCurColumnPos(24)
			C_cur_cd2					= iCurColumnPos(25)
			C_item_group_cd2			= iCurColumnPos(26)
			C_item_group_nm2			= iCurColumnPos(27)	
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

	arrParam(0) = "�����˾�"					' �˾� ��Ī 
	arrParam(1) = "B_PLANT"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"						' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"						' Field��(0)
    arrField(1) = "PLANT_NM"						' Field��(1)
    
    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
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
	
	arrField(0) = 1 							'ITEM_CD
	arrField(1) = 2 							'ITEM_NM
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

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
	arrParam(1) = frm1.txtFromDt.Text
	arrParam(2) = frm1.txtToDt.Text
	arrParam(3) = "OP"
	arrParam(4) = "OP"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "ǰ��׷�"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "ǰ��׷�"
	arrHeader(1) = "ǰ��׷��"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval strCode, Byval iWhere)
    
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
		
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtFromDt.Text
	arrParam(4) = frm1.txtToDt.Text	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
    	
End Function

'------------------------------------------  OpenOprRef()  -------------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.vspdData2.MaxRowS < 1 Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	frm1.vspdData2.Row =frm1.vspdData2.ActiveRow
	frm1.vspdData2.Col = C_prodt_order_no2
                
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.vspdData2.Text)			'��: ��ȸ ���� ����Ÿ 

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function


'------------------------------------------  OpenPartRef()  -------------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPartRef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4311RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4311RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	
   	With frm1.vspdData2
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_prodt_order_no2
		arrParam(1) = .Text
	End With

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function



'------------------------------------------  OpenProdRef()  -------------------------------------------------
'	Name : OpenProdRef()
'	Description : Production Reference
'---------------------------------------------------------------------------------------------------------
Function OpenProdRef()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)	'��: ��ȸ ���� ����Ÿ 

   	With frm1.vspdData2
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_prodt_order_no2
		arrParam(1) = .Text
	End With
	
	iCalledAspName = AskPRAspName("P4411RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4411RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  -------------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Issue Reference
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	
    With frm1.vspdData2
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_prodt_order_no2
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function LookUpItemByPlant(strCode)
	Dim strVal
    
    With frm1
    
    strVal = BIZ_PGM_ITEM_ID & "?txtCode=" & strCode						'��: ���� �� 
    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
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
'   Event Name : vspdData1_Change
'   Event Desc :
'==========================================================================================

Sub vspdData1_Change(ByVal Col , ByVal Row )

    
End Sub


'==========================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub


'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData1 

    If Row >= NewRow Then
        Exit Sub
    End If

	 '----------  Coding part  -------------------------------------------------------------

    End With

End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
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
 		
 		frm1.vspddata2.MaxRows = 0
    
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
 		
		lgOldRow1 = frm1.vspdData1.Row
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow1 <> Row Then
 			frm1.vspdData1.Row = Row
 			
 			frm1.vspddata2.MaxRows = 0
    
			If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
 			
			lgOldRow1 = Row
			
		End If	
	 	'------ Developer Coding part (End)
	
 	End If

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
    
     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
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
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("P")
		Frm1.txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("P")
		Frm1.txtToDt.Focus
    End If
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
 
'========================================================================================
' Function Name : vspdData1_DblClick1
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
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.ReOrderingSpreadData
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

    ggoSpread.Source = frm1.vspdData1										'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then									'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")				'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	If ValidDateCheck(frm1.txtFromDt, frm1.txtTODt) = False Then Exit Function
    	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables
					   														'��: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If															'��: Query db data

    FncQuery = True															'��: Processing is OK
    
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
    
    FncSave = False                                           '��: Processing is NG
    
    Err.Clear                                                 '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '��: Check required field(Multi area)
       Exit Function
    End If
        
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function				                                  '��: Save db data
    
    FncSave = True                                            '��: Processing is OK
    
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
    ggoSpread.Source = frm1.vspdData1	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
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

Function FncPrint()                                                       '��: Protect system from crashing
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
	
    ggoSpread.Source = frm1.vspdData1
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
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

    Err.Clear
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
	    strVal = strVal & "&txtInfNo=" & Trim(frm1.hInfNo.value)
	    strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)
	    strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
	    strVal = strVal & "&txtFromDt=" & Trim(frm1.hFromDt.value)
	    strVal = strVal & "&txtToDt=" & Trim(frm1.hToDt.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&rdoApplyFlag=" & Trim(frm1.hApplyFlag.value)		
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData1.MaxRows
	Else
	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	    strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)
	    strVal = strVal & "&txtInfNo=" & Trim(frm1.txtInfNo.value)
	    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtFromDt=" & frm1.txtFromDt.text
		strVal = strVal & "&txtToDt=" & frm1.txtToDt.text
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		If frm1.rdoApplyFlag1.checked = True Then
			strVal = strVal & "&rdoApplyFlag=" & frm1.rdoApplyFlag1.value 
		ElseIf frm1.rdoApplyFlag2.checked = True Then
			strVal = strVal & "&rdoApplyFlag=" & frm1.rdoApplyFlag2.value 
		Else 
			strVal = strVal & "&rdoApplyFlag=" & frm1.rdoApplyFlag3.value 
		End If
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData1.MaxRows
		
	End If    

    Call RunMyBizASP(MyBizASP, strVal)
	
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(Byval LngMaxRow)

	Call SetToolBar("11000000000111")
	
	Call InitData(LngMaxRow)
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
		
		lgOldRow1 = 1
		
    End If
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
	
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
'========================================================================================
Function DbQuery2() 
    
    DbQuery2 = False                                                         			'��: Processing is NG
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    frm1.vspdData1.Col = C_pop_inf_key
    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow

	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&txtInfNo=" & Trim(frm1.vspdData1.Text)

    Call RunMyBizASP(MyBizASP, strVal)									

    DbQuery2 = True                                                     

End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dbsave = False
	
    DbSave = True                                                           '��: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 

End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
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
    '----------  Coding part  -------------------------------------------------------------


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
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>POP�ݿ������ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPartRef()">��ǰ����</A> | <A href="vbscript:OpenOprRef()">��������</A> | <A href="vbscript:OpenProdRef()">��������</A> | <A href="vbscript:OpenRcptRef()">�԰���</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�ݿ���</TD>
									<TD CLASS="TD6">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������" id=fpDateTime1></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������" id=OBJECT1></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ݿ�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSchedIssueFlg" tag="11" ID="rdoApplyFlag1" VALUE="Y"><LABEL FOR="rdoSchedIssueFlg">��</LABEL>
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSchedIssueFlg" tag="11" ID="rdoApplyFlag2" VALUE="N"><LABEL FOR="rdoSchedIssueFlg">�ƴϿ�</LABEL>
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSchedIssueFlg" tag="11" CHECKED ID="rdoApplyFlag3" VALUE="A"><LABEL FOR="rdoSchedIssueFlg">��ü</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>�Է¹�ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInfNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="�Է¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value,0"></TD>
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
							<TR HEIGHT="80%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 ID = "A" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="20%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 ID = "B" WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hApplyFlag" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hInfNo" TABINDEX = "-1" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
