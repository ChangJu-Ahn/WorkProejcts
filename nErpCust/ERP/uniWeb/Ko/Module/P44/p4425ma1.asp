<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4425ma1.asp
'*  4. Program Name         : ������������ȸ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003-02-19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
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

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "p4425mb1.asp"							'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY2_ID	= "p4425mb2.asp"							'��: �����Ͻ� ���� ASP�� 

'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

' Grid 1(vspdData1) - Order
Dim C_ProdtOrderNo1		'����������ȣ 
Dim C_ItemCd1			'ǰ�� 
Dim C_ItemNm1			'ǰ��� 
Dim C_Spec1				'�԰� 
Dim C_OrderQty1			'�������� �������� 
Dim C_OrderProdQty1		'�������� �������� 
Dim C_OrderGoodQty1		'�������� ��ǰ���� 
Dim C_OrderBadQty1		'�������� �ҷ����� 
Dim C_OrderRcptQty1		'���������԰���� 
Dim C_OrderUnit1		'�������� 
Dim C_PlanStartDt1		'���������� 
Dim C_PlanComptDt1		'�ϷΌ���� 
Dim C_TrackingNo1		'Tracking No	
Dim C_OrderStatus1		'���û��� 
Dim C_BaseQty1			'���ش��� �������� 
Dim C_BaseProdQty1		'���ش��� �������� 
Dim C_BaseGoodQty1		'���ش��� ��ǰ���� 
Dim C_BaseBadQty1		'���ش��� �ҷ����� 
Dim C_BaseRcptQty1		'���ش��� �԰���� 
Dim C_BaseUnit1			'���ش��� 
Dim C_ItemGroupCd1
Dim C_ItemGroupNm1

' Grid 2(vspdData2) - Result
Dim C_ReportDt2			'������ 
Dim C_OrderProdQty2		'�������� ������� 
Dim C_OrderGoodQty2		'�������� ��ǰ���� 
Dim C_OrderBadQty2		'�������� �ҷ����� 
Dim C_OrderRcptQty2		'�������� �԰���� 
Dim C_OrderUnit2		'�������� 
Dim C_BaseProdQty2		'���ش��� ������� 
Dim C_BaseGoodQty2		'���ش��� ��ǰ���� 
Dim C_BaseBadQty2		'���ش��� �ҷ����� 
Dim C_BaseRcptQty2		'���ش��� �԰���� 
Dim C_BaseUnit2			'���ش��� 
'Hidden
Dim C_ProdtOrderNo2	'����������ȣ 

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size�� ������ ���� 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2

Dim strDate
Dim iDBSYSDate

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
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
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
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtReportFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtReportToDt.text   = UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_ItemGroupNm1 + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit 	C_ProdtOrderNo1, "����������ȣ"	,18
			ggoSpread.SSSetEdit 	C_ItemCd1,       "ǰ��"			,18
			ggoSpread.SSSetEdit 	C_ItemNm1,       "ǰ���"		,25
			ggoSpread.SSSetEdit 	C_Spec1,		 "�԰�"			,25
			ggoSpread.SSSetFloat 	C_OrderQty1,	 "��������"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_OrderProdQty1, "��������"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_OrderGoodQty1, "��ǰ����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_OrderBadQty1,	 "�ҷ�����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_OrderRcptQty1, "�԰����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_OrderUnit1,	 "��������"		,10
			ggoSpread.SSSetEdit 	C_PlanStartDt1,  "����������"	,11, 2
			ggoSpread.SSSetEdit 	C_PlanComptDt1,	 "�ϷΌ����"	,11, 2
			ggoSpread.SSSetEdit 	C_TrackingNo1,	 "Tracking No"	,25
			ggoSpread.SSSetEdit 	C_OrderStatus1,	 "���û���"		,10
			ggoSpread.SSSetFloat 	C_BaseQty1,		 "��������"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_BaseProdQty1,	 "��������"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_BaseGoodQty1,	 "��ǰ����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_BaseBadQty1,	 "�ҷ�����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_BaseRcptQty1,	 "�԰����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_BaseUnit1,	 "���ش���"		,10
			ggoSpread.SSSetEdit 	C_ItemGroupCd1, "ǰ��׷�"		,15
			ggoSpread.SSSetEdit		C_ItemGroupNm1, "ǰ��׷��"		,30
			
			Call ggoSpread.MakePairsColumn(C_OrderQty1, C_OrderUnit1 )
			Call ggoSpread.MakePairsColumn(C_BaseQty1, C_BaseUnit1 )
			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("A")
			
			.ReDraw = true    
    
		End With
	
    End If
    
    If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData2 
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021225", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_ProdtOrderNo2 + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit 	C_ReportDt2,		"������"		,10, 2
			ggoSpread.SSSetFloat	C_OrderProdQty2,    "��������"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_OrderGoodQty2,	"��ǰ����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_OrderBadQty2,		"�ҷ�����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_OrderRcptQty2,	"�԰����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_OrderUnit2,		"��������"		,10
			ggoSpread.SSSetFloat	C_BaseProdQty2,     "��������"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BaseGoodQty2,		"��ǰ����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BaseBadQty2,		"�ҷ�����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_BaseRcptQty2,		"�԰����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_BaseUnit2,		"���ش���"		,10
			ggoSpread.SSSetEdit 	C_ProdtOrderNo2,	"����������ȣ"	,18
			
			Call ggoSpread.MakePairsColumn(C_OrderProdQty2, C_OrderUnit2 )
			Call ggoSpread.MakePairsColumn(C_BaseProdQty2, C_BaseUnit2 )
			
			Call ggoSpread.SSSetColHidden( C_ProdtOrderNo2, C_ProdtOrderNo2, True)			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true    
    
		End With
    End If
    
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	If pvSpdNo = "A" Then
		'--------------------------------
		'Grid 1
		'--------------------------------
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
		
	If pvSpdNo = "B" Then 
		'--------------------------------
		'Grid 2
		'--------------------------------
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If	
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'========================== 2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboOrderStatus, lgF0, lgF1, Chr(11))
    
    frm1.cboOrderStatus.value = ""
End Sub

'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Order Header
		C_ProdtOrderNo1			= 1
		C_ItemCd1				= 2
		C_ItemNm1				= 3
		C_Spec1					= 4
		C_OrderQty1				= 5
		C_OrderProdQty1			= 6
		C_OrderGoodQty1			= 7
		C_OrderBadQty1			= 8
		C_OrderRcptQty1			= 9
		C_OrderUnit1			= 10
		C_PlanStartDt1			= 11
		C_PlanComptDt1			= 12
		C_TrackingNo1			= 13
		C_OrderStatus1			= 14
		C_BaseQty1				= 15
		C_BaseProdQty1			= 16
		C_BaseGoodQty1			= 17
		C_BaseBadQty1			= 18
		C_BaseRcptQty1			= 19
		C_BaseUnit1				= 20
		C_ItemGroupCd1			= 21
		C_ItemGroupNm1			= 22
	End If	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Result
		C_ReportDt2				= 1
		C_OrderProdQty2			= 2
		C_OrderGoodQty2			= 3
		C_OrderBadQty2			= 4
		C_OrderRcptQty2			= 5
		C_OrderUnit2			= 6
		C_BaseProdQty2			= 7
		C_BaseGoodQty2			= 8
		C_BaseBadQty2			= 9
		C_BaseRcptQty2			= 10
		C_BaseUnit2				= 11
		C_ProdtOrderNo2			= 12
	End If	

End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData1
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_ProdtOrderNo1			= iCurColumnPos(1)
			C_ItemCd1				= iCurColumnPos(2)
			C_ItemNm1				= iCurColumnPos(3)
			C_Spec1					= iCurColumnPos(4)
			C_OrderQty1				= iCurColumnPos(5)
			C_OrderProdQty1			= iCurColumnPos(6)
			C_OrderGoodQty1			= iCurColumnPos(7)
			C_OrderBadQty1			= iCurColumnPos(8)
			C_OrderRcptQty1			= iCurColumnPos(9)
			C_OrderUnit1			= iCurColumnPos(10)
			C_PlanStartDt1			= iCurColumnPos(11)
			C_PlanComptDt1			= iCurColumnPos(12)
			C_TrackingNo1			= iCurColumnPos(13)
			C_OrderStatus1			= iCurColumnPos(14)
			C_BaseQty1				= iCurColumnPos(15)
			C_BaseProdQty1			= iCurColumnPos(16)
			C_BaseGoodQty1			= iCurColumnPos(17)
			C_BaseBadQty1			= iCurColumnPos(18)
			C_BaseRcptQty1			= iCurColumnPos(19)
			C_BaseUnit1				= iCurColumnPos(20)
			C_ItemGroupCd1			= iCurColumnPos(21)
			C_ItemGroupNm1			= iCurColumnPos(22)
	
		Case "B"
		
			ggoSpread.Source = frm1.vspdData2
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_ReportDt2				= iCurColumnPos(1)
			C_OrderProdQty2			= iCurColumnPos(2)
			C_OrderGoodQty2			= iCurColumnPos(3)
			C_OrderBadQty2			= iCurColumnPos(4)
			C_OrderRcptQty2			= iCurColumnPos(5)
			C_OrderUnit2			= iCurColumnPos(6)
			C_BaseProdQty2			= iCurColumnPos(7)
			C_BaseGoodQty2			= iCurColumnPos(8)
			C_BaseBadQty2			= iCurColumnPos(9)
			C_BaseRcptQty2			= iCurColumnPos(10)
			C_BaseUnit2				= iCurColumnPos(11)
			C_ProdtOrderNo2			= iCurColumnPos(12)
			
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
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  ---------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If Trim(frm1.txtPlantCd.value) = "" Then
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""									'ProdFromDt
	arrParam(2) = ""									'ProdToDt
	arrParam(3) = Trim(frm1.cboOrderStatus.value)		'From Status
	arrParam(4) = Trim(frm1.cboOrderStatus.value)		'To Status
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = ""									'cboOrderType
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -----------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
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
	arrParam(1) = Trim(frm1.txtItemCd.value)    ' Item Code						
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenItemGroup()  -------------------------------------------------
'	Name : OpenItemGroup()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
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

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
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
	arrParam(3) = frm1.txtReportFromDt.Text
	arrParam(4) = frm1.txtReportToDt.Text	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
    
End Function

'------------------------------------------  OpenShiftCd()  ----------------------------------------------
'	Name : OpenShiftCd()
'	Description : Shift Popup
'---------------------------------------------------------------------------------------------------------
Function OpenShiftCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	 If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Shift �˾�"											' �˾� ��Ī 
	arrParam(1) = "P_SHIFT_HEADER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtShiftCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")	' Where Condition
	arrParam(5) = "Shift"												' TextBox ��Ī 
	 
    arrField(0) = "SHIFT_CD"												' Field��(0)
    arrField(1) = "DESCRIPTION"												' Field��(1)
    
    arrHeader(0) = "Shift"												' Header��(0)
    arrHeader(1) = "Shift ��"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetShiftCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtShiftCd.focus
		
End Function


'------------------------------------------  OpenPartRef()  ----------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPartRef()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
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
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo1
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
	End If	
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenOprRef()  -----------------------------------------------
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
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo1
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  ----------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Receipt Reference
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
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
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo1
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenConsumRef()  --------------------------------------------
'	Name : OpenConsumRef()
'	Description : Part Consumption Reference
'---------------------------------------------------------------------------------------------------------
Function OpenConsumRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
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
	
	iCalledAspName = AskPRAspName("P4412RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4412RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo1
		arrParam(1) = Trim(frm1.vspdData1.Text)				'��: ��ȸ ���� ����Ÿ 
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
'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus()		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value	= arrRet(0)
    frm1.txtProdOrderNo.focus()
End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
	frm1.txtItemCd.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItemCd.focus()
End Function

'------------------------------------------  SetItemInfo()  -------------------------------------------
'	Name : SetItemGroup()
'	Description : Item Group Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
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
	frm1.txtTrackingNo.focus()
End Function 

'------------------------------------------  SetShiftCd()  -------------------------------------------------
'	Name : SetShiftCd()
'	Description : Condition Shift Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetShiftCd(byval arrRet)
	frm1.txtShiftCd.Value    = arrRet(0)
	frm1.txtShiftCd.focus()
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

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitSpreadSheet("*")                                               '��: Setup the Spread sheet
   
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call InitComboBox
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

'=======================================================================================================
'   Event Name : txtReportFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportFromDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportFromDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtReportFromDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtReportToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtReportToDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtReportFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtReportFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtReportToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtReportToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_GotFocus
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
    
    If CheckRunningBizProcess = True Then							'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
        
        lgOldRow = Row
		If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
   
    Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then		
			frm1.vspdData1.Col = 1
			frm1.vspdData1.Row = row
		
			lgOldRow = Row
		
			If DbDtlQuery(Row) = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If		
	 	'------ Developer Coding part (End)
	
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

    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
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
    Else
        
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

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'==========================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

End Sub

'==========================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )

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
    End With
End Sub


'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If frm1.vspdData1.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If
	
	Call vspdData1_Click(NewCol, NewRow)
End Sub

'==========================================================================================
'   Event Name : vspdData2_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData2 
		If Row >= NewRow Then
		    Exit Sub
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

    FncQuery = False														'��: Processing is NG
    Err.Clear																'��: Protect system from crashing

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
	If ValidDateCheck(frm1.txtReportFromDt, frm1.txtReportToDt) = False Then Exit Function
		
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables														'��: Initializes local global variables
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'��: Query db data
	End If
	
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
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
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
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
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
	FncExit = True
End Function

'******************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  **************************
'	���� : 
'**************************************************************************************** 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal
       
    DbQuery = False

	Call LayerShowHide(1)
    
	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.hPlantCd.value))			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtReportFromDt="	& Trim(.hReportFromDt.value)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtReportToDt="		& Trim(.hReportToDt.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdOrderNo="	& UCase(Trim(.hProdOrderNo.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.hItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtTrackingNo="		& UCase(Trim(.hTrackingNo.value))
			strVal = strVal & "&txtShiftCd="		& UCase(Trim(.hShiftCd.value))		
			strVal = strVal & "&cboOrderStatus="	& UCase(Trim(.hOrderStatus.value))		
			strVal = strVal & "&txtItemGroupCd="	& UCase(Trim(.hItemGroupCd.value))
			strVal = strVal & "&lgStrPrevKey1="		& lgStrPrevKey1							'ProdOrderNo
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.txtPlantCd.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtReportFromDt="	& Trim(.txtReportFromDt.Text)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtReportToDt="		& Trim(.txtReportToDt.Text)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdOrderNo="	& UCase(Trim(.txtProdOrderNo.value))	'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.txtItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtTrackingNo="		& UCase(Trim(.txtTrackingNo.value))
			strVal = strVal & "&txtShiftCd="		& UCase(Trim(.txtShiftCd.value))		
			strVal = strVal & "&cboOrderStatus="	& UCase(Trim(.cboOrderStatus.value))
			strVal = strVal & "&txtItemGroupCd="	& UCase(Trim(.txtItemGroupCd.value))
			strVal = strVal & "&lgStrPrevKey1="		& lgStrPrevKey1
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
		End If
	End With
    
    Call RunMyBizASP(MyBizASP, strVal)														'��: �����Ͻ� ASP �� ���� 
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11000000000111")														'��: ��ư ���� ���� 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		If DbDtlQuery(frm1.vspdData1.Row) = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	lgOldRow = 1
		
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByVal LngRow) 
    
    Dim strVal
       
    DbDtlQuery = False

	Call LayerShowHide(1)
    
	With frm1
	
		.vspdData2.MaxRows = 0
		
		.vspdData1.Row = LngRow
		.vspdData1.Col = C_ProdtOrderNo1
				
		strVal = BIZ_PGM_QRY2_ID & "?txtMode="	& parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd="		& UCase(Trim(.hPlantCd.value))			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtReportFromDt="	& Trim(.hReportFromDt.value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtReportToDt="		& Trim(.hReportToDt.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdOrderNo="	& UCase(Trim(.vspdData1.text))			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtShiftCd="		& UCase(Trim(.hShiftCd.value))		
	End With
    
    Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
    
    DbDtlQuery = True
    
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
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub 

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������������ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPartRef()">��ǰ����</A> | <A href="vbscript:OpenOprRef()">��������</A> | <A href="vbscript:OpenRcptRef()">�԰���</A> | <A href="vbscript:OpenConsumRef()">��ǰ�Һ񳻿�</A></TD>					
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD> 
									<TD CLASS="TD6">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtReportFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtReportToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���û���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderStatus" ALT="���û���" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>Shift</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShiftCd" SIZE=8 MAXLENGTH=2 tag="11xxxU" ALT="Shift"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShiftCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenShiftCd()"></TD>
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
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> name=vspdData1 Id = "A" HEIGHT="100%" width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> name=vspdData2 Id = "B" HEIGHT="100%" width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hReportFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hReportToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hShiftCd" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
