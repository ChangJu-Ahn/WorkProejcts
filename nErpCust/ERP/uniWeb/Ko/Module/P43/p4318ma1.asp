
<%@ LANGUAGE="VBSCRIPT" %> 
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4318ma1.asp
'*  4. Program Name         : Routing Information
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/11/20
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Kang Hyo Ku
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
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

Option Explicit                                                      

On Error Resume Next
Err.Clear
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY1_ID						= "p4318mb1.asp"		'��: Release Production Order Query ASP�� 
Const BIZ_PGM_QRY2_ID						= "p4318mb2.asp"		'��: Operation Query ASP�� 
Const BIZ_PGM_QRY3_ID						= "p4318mb3.asp"		'��: Reservation Query ASP�� 

'-------------------------------
' Column Constants : Spread 1 
'-------------------------------
Dim C_ItemCd			
Dim C_ItemNm			
Dim C_ReqQty			
Dim C_ReqUnit			
Dim C_ResvdQty		
Dim C_IssuedQty		
Dim C_ConsumedQty		
Dim C_RemainQty		
Dim C_ItemGroupCd
Dim C_ItemGroupNm
Dim C_ItemAcct

'-------------------------------
' Column Constants : Spread 2 
'-------------------------------
Dim C_WCCd
Dim C_WcNm
Dim C_ReqQty1
Dim C_ResvdQty1
Dim C_IssuedQty1
Dim C_ConsumedQty1
Dim C_RemainQty1

'-------------------------------
' Column Constants : Spread 3 
'-------------------------------
Dim C_ProdtOrderNo
Dim C_OprNo
Dim C_ReqDt
Dim C_ReqQty2
Dim C_ResvdQty2
Dim C_IssuedQty2
Dim C_ConsumedQty2
Dim C_RemainQty2
Dim C_TrackingNo
Dim C_OrderStatus

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevKey1
Dim lgStrPrevKey2		
Dim lgStrPrevKey3

Dim lgSortKey1
Dim lgSortKey2
Dim lgSortKey3

Dim lgOldRow1
Dim lgOldRow2

Dim iDBSYSDate
Dim LocSvrDate
Dim StartDate
Dim EndDate

	iDBSYSDate = "<%=GetSvrDate%>"			'��: DB�� ���� ��¥�� �޾ƿͼ� ���۳�¥�� ����Ѵ�.
	LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	StartDate = UNIDateAdd("D",-7,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
	EndDate = UNIDateAdd("D", 7,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop					 'Popup
Dim gSelframeFlg


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
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgOldRow1 = 0
	lgOldRow2 = 0
	lgSortKey1 = 1
	lgSortKey2 = 1
	lgSortKey3 = 1
	
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

    frm1.txtReqStartDt.Text = StartDate
	frm1.txtReqEndDt.Text = EndDate

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()   
   <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
   <% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
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
			ggoSpread.Spreadinit "V20021206", ,Parent.gAllowDragDropSpread
					
			.vspdData1.ReDraw = false
     		.vspdData1.MaxCols = C_ItemAcct + 1
			.vspdData1.MaxRows = 0
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_ItemCd,		"ǰ��", 18
			ggoSpread.SSSetEdit		C_ItemNm,		"ǰ���", 25
			ggoSpread.SSSetFloat	C_ReqQty,		"�ʿ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_ReqUnit,		"����", 7,,,3,2
			ggoSpread.SSSetFloat	C_ResvdQty,		"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_IssuedQty,	"������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ConsumedQty,	"�Һ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty,	"����������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ItemGroupCd,	"ǰ��׷�", 12
			ggoSpread.SSSetEdit		C_ItemGroupNm,	"ǰ��׷��", 30
			ggoSpread.SSSetEdit		C_ItemAcct,		"ǰ�����", 12
			
			Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("A")
			
			.vspdData1.ReDraw = True
		End If
		
		If pvSpdNo = "B" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 2 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20021106", ,Parent.gAllowDragDropSpread
			.vspdData2.ReDraw = false
 
    		.vspdData2.MaxCols = C_RemainQty1 + 1
			.vspdData2.MaxRows = 0

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit		C_WCCd,			"�۾���", 10
			ggoSpread.SSSetEdit		C_WCNm,			"�۾����", 20
			ggoSpread.SSSetFloat	C_ReqQty1,		"�ʿ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ResvdQty1,	"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_IssuedQty1,	"������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ConsumedQty1,	"�Һ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty1,	"����������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

			Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols, .vspdData2.MaxCols, True)

			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("B")
			
			.vspdData2.ReDraw = True
		End If
		
		If pvSpdNo = "C" Or pvSpdNo = "*" Then	
			'-------------------------------------------
			' Spread 3 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData3
			ggoSpread.Spreadinit "V20021106", ,Parent.gAllowDragDropSpread
			.vspdData3.ReDraw = false
			.vspdData3.MaxCols = C_OrderStatus + 1
			.vspdData3.MaxRows = 0
	
			Call GetSpreadColumnPos("C")
		
			ggoSpread.SSSetEdit		C_ProdtOrderNo,	"������ȣ", 18
			ggoSpread.SSSetEdit		C_OprNo,		"����", 7,,,3		
			ggoSpread.SSSetDate 	C_ReqDt,		"�ʿ���", 11, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_ReqQty2,		"�ʿ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ResvdQty2,	"�������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_IssuedQty2,	"������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ConsumedQty2,	"�Һ����",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty2,	"����������",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25
			ggoSpread.SSSetEdit		C_OrderStatus,	"���û���", 10
		
			Call ggoSpread.SSSetColHidden(.vspdData3.MaxCols, .vspdData3.MaxCols, True)
		
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("C")
			
			.vspdData3.ReDraw = True
		End If	
		
    End With
        
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		If pvSpdNo = "A" Then
			'-------------------------
			' Set Lock Prop :Spread 1 		
			'-------------------------
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "B" Then		
			'-------------------------
			' Set Lock Prop :Spread 2 		
			'-------------------------
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "C" Then
			'-------------------------
			' Set Lock Prop :Spread 3 		
			'-------------------------
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	

    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))

End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
    
    If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'grid1
		C_ItemCd			= 1
		C_ItemNm			= 2																	'��: Spread Sheet �� Columns �ε��� 
		C_ReqQty			= 3
		C_ReqUnit			= 4
		C_ResvdQty			= 5
		C_IssuedQty			= 6
		C_ConsumedQty		= 7
		C_RemainQty			= 8
		C_ItemGroupCd		= 9
		C_ItemGroupNm		= 10
		C_ItemAcct			= 11
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then	
		'grid2
		C_WCCd				= 1
		C_WcNm				= 2
		C_ReqQty1			= 3
		C_ResvdQty1			= 4
		C_IssuedQty1		= 5
		C_ConsumedQty1		= 6
		C_RemainQty1		= 7
	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then	
		'grid3
		C_ProdtOrderNo		= 1
		C_OprNo				= 2
		C_ReqDt				= 3
		C_ReqQty2			= 4
		C_ResvdQty2			= 5
		C_IssuedQty2		= 6
		C_ConsumedQty2		= 7
		C_RemainQty2		= 8
		C_TrackingNo		= 9
		C_OrderStatus		= 10

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
					
			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)																	'��: Spread Sheet �� Columns �ε��� 
			C_ReqQty			= iCurColumnPos(3)
			C_ReqUnit			= iCurColumnPos(4)
			C_ResvdQty			= iCurColumnPos(5)
			C_IssuedQty			= iCurColumnPos(6)
			C_ConsumedQty		= iCurColumnPos(7)
			C_RemainQty			= iCurColumnPos(8)
			C_ItemGroupCd		= iCurColumnPos(9)
			C_ItemGroupNm		= iCurColumnPos(10)
			C_ItemAcct			= iCurColumnPos(11)
			
		Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_WCCd				= iCurColumnPos(1)
			C_WcNm				= iCurColumnPos(2)
			C_ReqQty1			= iCurColumnPos(3)
			C_ResvdQty1			= iCurColumnPos(4)
			C_IssuedQty1		= iCurColumnPos(5)
			C_ConsumedQty1		= iCurColumnPos(6)
			C_RemainQty1		= iCurColumnPos(7)
					
		Case "C"
		 	ggoSpread.Source = frm1.vspdData3
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ProdtOrderNo		= iCurColumnPos(1)
			C_OprNo				= iCurColumnPos(2)
			C_ReqDt				= iCurColumnPos(3)
			C_ReqQty2			= iCurColumnPos(4)
			C_ResvdQty2			= iCurColumnPos(5)
			C_IssuedQty2		= iCurColumnPos(6)
			C_ConsumedQty2		= iCurColumnPos(7)
			C_RemainQty2		= iCurColumnPos(8)
			C_TrackingNo		= iCurColumnPos(9)
			C_OrderStatus		= iCurColumnPos(10)
					
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
		Call displaymsgbox("971012","X" , "����","X")
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
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ���: From To�� �Է��� �� 
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
		Call displaymsgbox("971012","X" , "����","X")
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

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
    
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True  Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
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
	arrParam(3) = ""
	arrParam(4) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
		
End Function

'------------------------------------------  OpenWcCd()  ------------------------------------------------
'	Name : OpenWcCd()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""	'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ""	' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
	
End Function

'------------------------------------------  OpenItemGroup()  --------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "ǰ��׷�"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    arrField(3) = "LEAF_FLG"	
    arrField(4) = "UPPER_ITEM_GROUP_CD"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

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

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function    

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemGroup()  ---------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value	= arrRet(0)		
	frm1.txtItemGroupNm.value   = arrRet(1)
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

	Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
	Call InitSpreadSheet("*")															'��: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal
	Call InitVariables																'��: Initializes local global variables
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11000000000011")
		
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
'	Name : txtReqStartDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtReqEndDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================

Sub txtReqStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqStartDt.Focus
    End If
End Sub

Sub txtReqEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqEndDt.Focus
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	
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
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey1
            lgSortKey1 = 1
        End If
        
        lgOldRow2 = 0
			
		lgStrPrevKey2 = ""
		lgStrPrevKey3 = ""
		frm1.vspdData2.MaxRows = 0
		frm1.vspdData3.MaxRows = 0
			
		frm1.vspdData1.Col = C_ItemCd
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
			
		frm1.KeyItemCd.value =  Trim(frm1.vspdData1.Text)
			
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
        
	Else
		If lgOldRow1 <> Row Then
			
			lgOldRow2 = 0
			
			lgStrPrevKey2 = ""
			lgStrPrevKey3 = ""
			frm1.vspdData2.MaxRows = 0
			frm1.vspdData3.MaxRows = 0
			
			frm1.vspdData1.Col = C_ItemCd
			frm1.vspdData1.Row = Row
			
			frm1.KeyItemCd.value =  Trim(frm1.vspdData1.Text)
			
			If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
					
			lgOldRow1 = Row
		End If
	End If	
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================

Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	
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
            ggoSpread.SSSort Col
            lgSortKey2 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey2
            lgSortKey2 = 1
        End If
        
		lgStrPrevKey3 = ""
			
		frm1.vspdData3.MaxRows = 0

		frm1.vspdData2.Row = frm1.vspdData2.ActiveRow		
		frm1.vspdData2.Col = C_WcCd
		frm1.KeyWcCd.value  = Trim(frm1.vspdData2.Text) 
			
		lgOldRow2 = frm1.vspdData2.Row
			
		If DbQuery3 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If	    
        
    ELSE
		
		If lgOldRow2 <> Row Then
			
			lgStrPrevKey3 = ""
			
			frm1.vspdData3.MaxRows = 0

			frm1.vspdData2.Row = Row		
			frm1.vspdData2.Col = C_WcCd
			frm1.KeyWcCd.value  = Trim(frm1.vspdData2.Text) 
			
			lgOldRow2 = Row
			
			If DbQuery3 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
			
		End If
		
	End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData3_Click
'   Event Desc :
'==========================================================================================
Sub vspdData3_Click(ByVal Col , ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	
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
            ggoSpread.SSSort Col
            lgSortKey3 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey3
            lgSortKey3 = 1
        End If
        Exit Sub
    End If
	
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

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData3_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData3_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP3C" Then
       gMouseClickStatus = "SP3CR"
    End If
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

Sub vspdData3_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("C")
End Sub

'==========================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc :
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
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1 ,NewTop) Then
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
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2 ,NewTop) Then
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
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3 ,NewTop) Then
		If lgStrPrevKey3 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery3 = False Then	
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
    'Erase contents area
    '-----------------------
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If ValidDateCheck(frm1.txtReqStartDt, frm1.txtReqEndDt) = False Then Exit Function	
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
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
    If DbQuery = False Then Exit Function																'��: Query db data
       
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
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&cboItemAcct=" & Trim(frm1.hItemAcct.value)
		strVal = strVal & "&txtReqStartDt=" & Trim(frm1.hReqStartDt.value)
		strVal = strVal & "&txtReqEndDt=" & Trim(frm1.hReqEndDt.value)
	Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtWcCd=" & Trim(frm1.txtWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&cboItemAcct=" & Trim(frm1.cboItemAcct.value)
		strVal = strVal & "&txtReqStartDt=" & Trim(frm1.txtReqStartDt.Text)
		strVal = strVal & "&txtReqEndDt=" & Trim(frm1.txtReqEndDt.Text)
	End If    

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
    
    ggoSpread.Source = frm1.vspdData1

	frm1.vspdData1.ReDraw = False
	
	frm1.vspdData1.ReDraw = True
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		frm1.vspdData1.Col = C_ItemCd
		frm1.vspdData1.Row = 1
		frm1.KeyItemCd.value = Trim(frm1.vspdData1.Text)
			
		lgOldRow1 = 1
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
		
    End If

    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field
    
    Call SetToolBar("11000000000111")
    
    lgBlnFlgChgValue = False
    
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
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.KeyItemCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtWcCd=" & Trim(frm1.hWcCd.value)					'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtReqStartDt=" & Trim(frm1.hReqStartDt.value)
	strVal = strVal & "&txtReqEndDt=" & Trim(frm1.hReqEndDt.value)
    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery2 = True                                                          	

End Function

'========================================================================================
' Function Name : DbQuery2Ok
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
'========================================================================================

Function DbQuery2Ok() 
	
	frm1.vspdData2.Col = C_WcCd
	frm1.vspdData2.Row = 1
	lgOldRow2 = 1
	
	frm1.KeyWcCd.value = Trim(frm1.vspdData2.Text)
	
	Call DbQuery3()
	
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
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.KeyItemCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtWcCd=" & Trim(frm1.KeyWcCd.value)				'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtReqStartDt=" & Trim(frm1.hReqStartDt.value)
	strVal = strVal & "&txtReqEndDt=" & Trim(frm1.hReqEndDt.value)	
    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery3 = True                                                          	

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 

End Function    

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
	Call ggoSpread.ReOrderingSpreadData()
End Sub 


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���������Ȳ��ȸ</font></td>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="����"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�ʿ���</TD>
								    <TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4318ma1_fpDateTime1_txtReqStartDt.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/p4318ma1_fpDateTime1_txtReqEndDt.js'></script>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>							
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWCCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()"> <INPUT TYPE=TEXT NAME="txtWCNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="ǰ�����" STYLE="Width: 160px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p4318ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4318ma1_B_vspdData2.js'></script>
								</TD>
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4318ma1_C_vspdData3.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="KeyWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqStartDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqEndDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
