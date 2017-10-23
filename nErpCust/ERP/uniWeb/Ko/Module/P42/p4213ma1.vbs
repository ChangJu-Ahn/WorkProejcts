
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY1_ID						= "p4213mb1.asp"		'��: Release Production Order Query ASP�� 
Const BIZ_PGM_QRY2_ID						= "p4213mb2.asp"		'��: Operation Query ASP�� 
Const BIZ_PGM_QRY3_ID						= "p4213mb3.asp"		'��: Reservation Query ASP�� 
Const BIZ_PGM_SAVE_ID						= "p4213mb4.asp"		'��: Cancel Release Production Order ASP�� 

'-------------------------------
' Column Constants : Spread 1 
'-------------------------------
Dim  C_Select			'= 1
Dim  C_OrdNo			'= 2																	'��: Spread Sheet �� Columns �ε��� 
Dim  C_ItemCd			'= 3
Dim  C_ItemNm			'= 4
Dim	 C_Spec				'= 5
Dim  C_RoutNo			'= 6
Dim  C_PlanStartDt		'= 7
Dim  C_PlanEndDt		'= 8
Dim  C_OrdQty			'= 9
Dim  C_OrdUnit			'= 10
Dim  C_TrackingNo		'= 11
Dim  C_OrderType		'= 12
Dim  C_OrderTypeDesc	'= 13
Dim C_ItemGroupCd		'= 14
Dim C_ItemGroupNm		'= 15

'-------------------------------
' Column Constants : Spread 2 
'-------------------------------

Dim  C_OprNo			'= 1																	'��: Spread Sheet �� Columns �ε��� 
Dim  C_WCCd				'= 2
Dim  C_WcNm				'= 3
Dim  C_JobCd			'= 4
Dim  C_JobNm			'= 5
Dim  C_PlanStartDt2		'= 6
Dim  C_PlanEndDt2		'= 7
Dim  C_InsideFlg		'= 8

'-------------------------------
' Column Constants : Spread 3 
'-------------------------------

Dim  C_ChildItemCd		'= 1
Dim  C_ChildItemNm		'= 2
Dim	 C_Spec3			'= 3
Dim  C_ReqDt			'= 4
Dim  C_ReqQty			'= 5
Dim  C_ResvUnit			'= 6
Dim  C_IssuedQty		'= 7


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgStrPrevKey2		
Dim lgStrPrevKey3

Dim lgSortKey1
Dim lgSortKey2
Dim lgSortKey3

Dim lgOldRow1
Dim lgOldRow2
		
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop					 'Popup
Dim gSelframeFlg

Dim lgButtonSelection
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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgOldRow1 = 0
	lgOldRow2 = 0
	lgSortKey1 = 1
	lgSortKey2 = 1
	lgSortKey3 = 1
	lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "��ü����"

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
	frm1.txtProdFromDt.text = StartDate
    frm1.txtProdToDt.text   = EndDate
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "��ü����"
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
    
    Call InitSpreadPosVariables(pvSpdNo)
        
    With frm1
		If pvSpdNo = "A" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 1 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData1
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			.vspdData1.ReDraw = False
    
			.vspdData1.MaxCols = C_ItemGroupNm + 1
			.vspdData1.MaxRows = 0
		
			Call GetSpreadColumnPos("A")
					
			ggoSpread.SSSetCheck		C_Select,		"",2,,,1
			ggoSpread.SSSetEdit			C_OrdNo,		"������ȣ", 18
			ggoSpread.SSSetEdit			C_ItemCd,		"ǰ��", 18
			ggoSpread.SSSetEdit			C_ItemNm,		"ǰ���", 25
			ggoSpread.SSSetEdit			C_Spec,			"�԰�", 25
			ggoSpread.SSSetEdit			C_RoutNo,		"�����", 10
			ggoSpread.SSSetDate 		C_PlanStartDt,	"����������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 		C_PlanEndDt,	"�ϷΌ����", 11, 2, parent.gDateFormat
			ggoSpread.SSSetFloat		C_OrdQty,		"��������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 		C_OrdUnit,		"����", 8,,,3,2
			ggoSpread.SSSetEdit			C_TrackingNo,	"Tracking No.", 25
			ggoSpread.SSSetCombo		C_OrderType,	"���ñ���", 10
			ggoSpread.SSSetCombo		C_OrderTypeDesc,"���ñ���", 10		
			ggoSpread.SSSetEdit		 	C_ItemGroupCd,	"ǰ��׷�",	15
			ggoSpread.SSSetEdit			C_ItemGroupNm,	"ǰ��׷��", 30
		
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden(C_OrderType, C_OrderType, True)
 			Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
		
			ggoSpread.SSSetSplit2(3)
			
			Call SetSpreadLock("A")
				
			.vspdData1.ReDraw = True
		End If
			
		If pvSpdNo = "B" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 2 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		
			.vspdData2.ReDraw = False
		
			.vspdData2.MaxCols = C_InsideFlg + 1
			.vspdData2.MaxRows = 0
		
			Call GetSpreadColumnPos("B")
		
			ggoSpread.SSSetEdit		C_OprNo,		"����", 10,,,3
			ggoSpread.SSSetEdit		C_WCCd,			"�۾���", 10
			ggoSpread.SSSetEdit		C_WCNm,			"�۾����", 20
			ggoSpread.SSSetEdit		C_JobCd,		"�۾�", 10
			ggoSpread.SSSetEdit		C_JobNm,		"�۾���", 20
			ggoSpread.SSSetDate 	C_PlanStartDt2,	"����������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_PlanEndDt2,	"�ϷΌ����", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_InsideFlg,	"�系/��", 8
		
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( .vspdData2.MaxCols, .vspdData2.MaxCols, True)
 					
			ggoSpread.SSSetSplit2(2)
		
			Call SetSpreadLock("B")
		
			.vspdData2.Redraw = False
		End If
		
		If pvSpdNo = "C" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 3 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData3
			ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		
			.vspdData3.ReDraw = False
		
			.vspdData3.MaxCols = C_IssuedQty + 1
			.vspdData3.MaxRows = 0
		
			Call GetSpreadColumnPos("C")
		
			ggoSpread.SSSetEdit		C_ChildItemCd,	"��ǰ��", 18
			ggoSpread.SSSetEdit		C_ChildItemNm,	"��ǰ���", 25
			ggoSpread.SSSetEdit		C_Spec3,		"�԰�", 25
			ggoSpread.SSSetDate 	C_ReqDt,		"�ʿ���", 11, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_ReqQty,		"�ʿ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_ResvUnit,		"����", 8,,,3,2
			ggoSpread.SSSetFloat	C_IssuedQty,	"������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		
			'Call ggoSpread.MakePairsColumn(,)
 			Call ggoSpread.SSSetColHidden( .vspdData3.MaxCols, .vspdData3.MaxCols, True)
		
			ggoSpread.SSSetSplit2(1)
		
			.vspdData3.ReDraw = True
			
			Call SetSpreadLock("C")
			
		End If
		
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		Select Case pvSpdNo
			Case "A"
				'-------------------------
				' Set Lock Prop :Spread 1 		
				'-------------------------	
				.vspdData1.ReDraw = False
				ggoSpread.Source = .vspdData1
				ggoSpread.SpreadLock -1, -1	' Set Lock Property : Spread 1
				ggoSpread.spreadUnLock C_Select, -1, C_Select
				.vspdData1.ReDraw = True
			Case "B"
				'-------------------------
				' Set Lock Prop :Spread 2 		
				'-------------------------
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadLockWithOddEvenRowColor()
			Case "C"
				'-------------------------
				' Set Lock Prop :Spread 3 		
				'-------------------------
				.vspdData3.ReDraw = False
				ggoSpread.Source = .vspdData3
				ggoSpread.SpreadLock -1, -1							' Set Lock Property : Spread 3
				.vspdData3.ReDraw = True
		End Select
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
End Sub

'==========================================  2.2.6 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderTypeDesc
   
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData1
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_OrderType
			intIndex = .value
			.Col = C_OrderTypeDesc
			.value = intindex
		Next
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 1 
		'-------------------------------
		C_Select			= 1
		C_OrdNo				= 2																	'��: Spread Sheet �� Columns �ε��� 
		C_ItemCd			= 3
		C_Spec				= 4
		C_ItemNm			= 5
		C_RoutNo			= 6
		C_PlanStartDt		= 7
		C_PlanEndDt			= 8
		C_OrdQty			= 9
		C_OrdUnit			= 10
		C_TrackingNo		= 11
		C_OrderType			= 12
		C_OrderTypeDesc		= 13
		C_ItemGroupCd		= 14
		C_ItemGroupNm		= 15
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 2 
		'-------------------------------
		C_OprNo				= 1																	'��: Spread Sheet �� Columns �ε��� 
		C_WCCd				= 2
		C_WcNm				= 3
		C_JobCd				= 4
		C_JobNm				= 5
		C_PlanStartDt2		= 6
		C_PlanEndDt2		= 7
		C_InsideFlg			= 8
	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		'-------------------------------
		' Column Constants : Spread 3 
		'-------------------------------
		C_ChildItemCd		= 1
		C_ChildItemNm		= 2
		C_Spec				= 3
		C_ReqDt				= 4
		C_ReqQty			= 5
		C_ResvUnit			= 6
		C_IssuedQty			= 7
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
			
			C_Select			= iCurColumnPos(1)
			C_OrdNo				= iCurColumnPos(2)																	'��: Spread Sheet �� Columns �ε��� 
			C_ItemCd			= iCurColumnPos(3)
			C_ItemNm			= iCurColumnPos(4)
			C_Spec				= iCurColumnPos(5)
			C_RoutNo			= iCurColumnPos(6)
			C_PlanStartDt		= iCurColumnPos(7)
			C_PlanEndDt			= iCurColumnPos(8)
			C_OrdQty			= iCurColumnPos(9)
			C_OrdUnit			= iCurColumnPos(10)
			C_TrackingNo		= iCurColumnPos(11)
			C_OrderType			= iCurColumnPos(12)
			C_OrderTypeDesc		= iCurColumnPos(13)
			C_ItemGroupCd		= iCurColumnPos(14)
			C_ItemGroupNm		= iCurColumnPos(15)
		
		Case "B"
			ggoSpread.Source = frm1.vspdData2 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
			C_OprNo				= iCurColumnPos(1)																	'��: Spread Sheet �� Columns �ε��� 
			C_WCCd				= iCurColumnPos(2)
			C_WcNm				= iCurColumnPos(3)
			C_JobCd				= iCurColumnPos(4)
			C_JobNm				= iCurColumnPos(5)
			C_PlanStartDt2		= iCurColumnPos(6)
			C_PlanEndDt2		= iCurColumnPos(7)
			C_InsideFlg			= iCurColumnPos(8)
		
		Case "C"
			ggoSpread.Source = frm1.vspdData3 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 				
			C_ChildItemCd		= iCurColumnPos(1)
			C_ChildItemNm		= iCurColumnPos(2)
			C_Spec3				= iCurColumnPos(3)
			C_ReqDt				= iCurColumnPos(4)
			C_ReqQty			= iCurColumnPos(5)
			C_ResvUnit			= iCurColumnPos(6)
			C_IssuedQty			= iCurColumnPos(7)
	
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

'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox �� 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
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
	arrParam(2) = "12!MO"							' Combo Set Data:"1029!MP" -- ǰ����� ������ ���ޱ���: From To�� �Է��� �� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
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

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdToDt.Text
	arrParam(3) = "RL"
	arrParam(4) = "RL"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value) 
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = Trim(frm1.cboOrderType.value)
		
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
Function OpenTrackingInfo()
	
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
	arrParam(3) = frm1.txtprodFromDt.Text
	arrParam(4) = frm1.txtprodToDt.Text	
	
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
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
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

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
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

Function btnAutoSel_onClick()

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "��ü����"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "��ü�������"
	End If

	Dim index,Count
	Dim strFlag
	
	frm1.vspdData1.ReDraw = false
	
	Count = frm1.vspdData1.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData1.Row = index
		frm1.vspdData1.Col = C_Select
		
		strFlag = frm1.vspdData1.Value
		
		If lgButtonSelection = "SELECT" Then
			frm1.vspdData1.Value = 1
			frm1.vspdData1.Col = 0 
			ggoSpread.UpdateRow Index
		Else
			frm1.vspdData1.Value = 0
			frm1.vspdData1.Col = 0 
			'ggoSpread.SSDeleteFlag Index
			frm1.vspdData1.Text=""
		End if

	Next 
	
	frm1.vspdData1.ReDraw = true

End Function

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtProdFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtProdToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
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
 		
 		lgOldRow2 = 0
				
		lgStrPrevKey2 = ""
		lgStrPrevKey3 = ""
		frm1.vspdData2.MaxRows = 0
		frm1.vspdData3.MaxRows = 0
		frm1.KeyOprNo.value = ""
				
		frm1.vspdData1.Col = C_OrdNo
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
				
		frm1.KeyProdOrdNo.value =  Trim(frm1.vspdData1.Text)
				
		IF DbQuery2 = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
		
		lgOldRow1 = frm1.vspdData1.Row 
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow1 <> Row Then
		
			lgOldRow2 = 0
				
			lgStrPrevKey2 = ""
			lgStrPrevKey3 = ""
			frm1.vspdData2.MaxRows = 0
			frm1.vspdData3.MaxRows = 0
			frm1.KeyOprNo.value = ""
				
			frm1.vspdData1.Col = C_OrdNo
			frm1.vspdData1.Row = Row
				
			frm1.KeyProdOrdNo.value =  Trim(frm1.vspdData1.Text)
				
			IF DbQuery2 = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
			lgOldRow1 = Row
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
 		
 		lgStrPrevKey3 = ""
				
		frm1.vspdData3.MaxRows = 0
				
		frm1.vspdData2.Col = C_OprNo
		frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
				
		frm1.KeyOprNo.value  = Trim(frm1.vspdData2.Text) 
				
		If DbQuery3 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
				
		lgOldRow2 = frm1.vspdData2.Row
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow2 <> Row Then
				
			lgStrPrevKey3 = ""
				
			frm1.vspdData3.MaxRows = 0
				
			frm1.vspdData2.Col = C_OprNo
			frm1.vspdData2.Row = Row
				
			frm1.KeyOprNo.value  = Trim(frm1.vspdData2.Text) 
				
			If DbQuery3 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
				
			lgOldRow2 = Row
		End If
	 	'------ Developer Coding part (End)
	
 	End If

End Sub

'==========================================================================================
'   Event Name : vspdData3_Click
'   Event Desc :
'==========================================================================================
Sub vspdData3_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP3C"
	
	Set gActiveSpdSheet = frm1.vspdData3
    
 	If frm1.vspdData3.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData3 
 		If lgSortKey3 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey3 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey3		'Sort in Descending
 			lgSortKey3 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
	
 	End If

End Sub


'==========================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc :
'==========================================================================================

Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData1
		
		ggoSpread.Source = frm1.vspdData1
		
		.Row = Row
		.Col = C_Select
		If ButtonDown = 1 Then
			ggoSpread.UpdateRow Row
		Else
			ggoSpread.SSDeleteFlag Row,Row
		End If			
		
	End With
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
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If 
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
			
		End If
    End if
    
End Sub
'==========================================================================================
'   Event Name : vspdData3_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then
		If lgStrPrevKey3 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery3 = False Then	
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
'   Event Name : vspdData3_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData3_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If

End Sub

'========================================================================================
' Function Name : vspdData1_DblClick
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
' Function Name : vspdData2_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
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
 
'========================================================================================
' Function Name : vspdData3_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData3_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData3.MaxRows = 0 Then
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
' Function Name : vspdData2_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData3_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	If NewCol = C_Select or Col = C_Select Then
		Cancel = True
		Exit Sub
	End If

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
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
' Function Name : vspdData3_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData3_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData3
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
    If gActiveSpdSheet = "A" Then Call InitSpreadComboBox()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.ReOrderingSpreadData
    If gActiveSpdSheet = "A" Then Call InitData(1)
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
	
	If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function	
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
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
		Call RestoreToolBar()
		Exit Function
	End If																	'��: Query db data
       
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
        
    FncSave = False												'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
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
	If frm1.vspdData1.MaxRows <= 0 Then Exit Function	

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.EditUndo            
	
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
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
    Call parent.FncExport(parent.C_SINGLEMULTI)									'��: ȭ�� ���� 
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
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                               '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
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
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : Spread 1 ��ȸ �� Scroll
'========================================================================================
Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtProdFromDt=" & Trim(frm1.hProdFromDt.value)
		strVal = strVal & "&txtProdToDt=" & Trim(frm1.hProdToDt.value)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.hOrderType.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)		
	Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtProdFromDt=" & Trim(frm1.txtProdFromDt.text)
		strVal = strVal & "&txtProdToDt=" & Trim(frm1.txtProdToDt.text)
		strVal = strVal & "&cboOrderType=" & Trim(frm1.cboOrderType.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
	End If    

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(Byval LngMaxRow)
    '-----------------------
    'Reset variables area
    '-----------------------
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		frm1.vspdData1.Col = C_OrdNo
		frm1.vspdData1.Row = 1
		frm1.KeyProdOrdNo.value = Trim(frm1.vspdData1.Text)
			
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Function
		End If	
		
		lgOldRow1 = 1
		
    End If

    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
	Call InitData(LngMaxRow)
	frm1.btnAutoSel.disabled = False
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field
    Call SetToolBar("11001001000111")
    
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 
'========================================================================================

Function DbQuery2() 
    
    DbQuery2 = False                                    
    
    Call LayerShowHide(1)
 
    Dim strVal

	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.KeyProdOrdNo.value)							
	
    Call RunMyBizASP(MyBizASP, strVal)											
	
    DbQuery2 = True                                                          	

End Function

'========================================================================================
' Function Name : DbQuery2Ok
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
'========================================================================================

Function DbQuery2Ok() 

	frm1.vspdData2.Col = C_OprNo
	frm1.vspdData2.Row = 1
	lgOldRow2 = 1
	
	frm1.KeyOprNo.value = Trim(frm1.vspdData2.Text)
	
	If DbQuery3 = False Then	
		Call RestoreToolBar()
		Exit Function
	End If	
	
	
	
End Function

'========================================================================================
' Function Name : DbQuery3
' Function Desc : Spread 2 Scroll 
'========================================================================================

Function DbQuery3() 
    
    DbQuery3 = False                                    
    
    Call LayerShowHide(1)
 
    Dim strVal
	
	strVal = BIZ_PGM_QRY3_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
	strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.KeyProdOrdNo.value)							
	strVal = strVal & "&txtOprNo=" & Trim(frm1.KeyOprNo.value)							
	
    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery3 = True                                                          	

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
	Dim lRow        
    Dim lGrpCnt     
    Dim strVal
    
    Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

    DbSave = False                                                          	'��: Processing is NG
    
    Call LayerShowHide(1)
    
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			

	iTmpCUBufferCount = -1
	
	strCUTotalvalLen = 0 

    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData1.MaxRows
    
        .vspdData1.Row = lRow
        .vspdData1.Col = 0

        Select Case .vspdData1.Text
                
            Case ggoSpread.UpdateFlag
				
				strVal = ""
				strVal = strVal & UCase(Trim(frm1.hPlantCd.value)) & iColSep			

				.vspdData1.Col = C_OrdNo			
				strVal = strVal & Trim(.vspdData1.Text) & iColSep
				strVal = strVal & lRow & iRowSep  
				
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
		
        End Select
                
    Next
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
	Call InitVariables
    
    frm1.vspdData1.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
    frm1.vspdData3.MaxRows = 0
	
	frm1.KeyOprNo.value = ""
	frm1.KeyProdOrdNo.value = ""
	
	Call RemovedivTextArea
	 
    Call MainQuery()

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
End Function
