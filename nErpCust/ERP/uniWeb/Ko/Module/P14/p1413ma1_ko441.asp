<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : P1413MA1.asp
'*  4. Program Name         : ��ǰ�� �ϰ���ü 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003-03-21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID			= "p1413mb1_ko441.asp"							'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID			= "p1413mb2_ko441.asp"							'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_LOOKUP_ECN_INFO	= "p1413mb3_ko441.asp"
Const BIZ_PGM_LOOKUP_ITEM_INFO	= "p1413mb4_ko441.asp"

'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
'Const C_SHEETMAXROWS = 50

' Grid 1(vspdData) - Operation
Dim C_Select		
Dim C_PrntItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_AcctNm
Dim C_ProcurTypeNm
Dim C_ChildSeq
Dim C_ChildItemCd
Dim C_ChildItemPopup
Dim C_ChildItemNm
Dim C_ChildSpec
Dim C_ChildAcctNm
Dim C_ChildProcurTypeNm
Dim C_ChildItemQty
Dim C_ChildUnit
Dim C_ChildUnitPopup
Dim C_PrntItemQty
Dim C_PrntUnit
Dim C_PrntUnitPopup
Dim C_SafetyLt
Dim C_LossRate
Dim C_SupplyType
Dim C_SupplyTypeNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_ECNNo
Dim C_ECNNoPopup
Dim C_ECNDesc
Dim C_ReasonCd
Dim C_ReasonCdPopup
Dim C_ReasonNm
Dim C_Remark

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							<%'Variable is for Dirty flag%>
Dim lgIntGrpCount								<%'Group View Size�� ������ ���� %>
Dim lgIntFlgMode								<%'Variable is for Operation Status%>

Dim lgStrPrevKey1
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey
Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

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
    lgSortKey = 1
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
	frm1.txtBomType.value = "1"
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,2) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        	frm1.txtPlantCd.value = lgPLCd
	End If
	'frm1.txtFromReqDt.text	= UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)
	'frm1.txtToReqDt.text	= UniConvDateAToB(UNIDateAdd ("D", 7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	'frm1.txtDoDate.text		= UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I","P","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	Call AppendNumberPlace("6","3","0")
	Call AppendNumberPlace("7","2","2")
	Call AppendNumberPlace("8","11","6")
	
	With frm1.vspdData 
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021122", ,Parent.gAllowDragDropSpread
				
		.ReDraw = false
				
		.MaxCols = C_Remark + 1    
		.MaxRows = 0    
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCheck	C_Select ,		"",2,,,1
		ggoSpread.SSSetEdit 	C_PrntItemCd,	"��ǰ��"	, 20
		ggoSpread.SSSetEdit 	C_ItemNm,       "��ǰ���"	, 30
		ggoSpread.SSSetEdit 	C_Spec,			"�԰�"		, 30
		ggoSpread.SSSetEdit 	C_AcctNm,		"ǰ�����"	, 10
		ggoSpread.SSSetEdit 	C_ProcurTypeNm,	"���ޱ���"	, 10
		ggoSpread.SSSetEdit		C_ChildSeq,		"����"		, 6
		ggoSpread.SSSetEdit 	C_ChildItemCd,	"��üǰ"	, 20,,,18,2
		ggoSpread.SSSetButton 	C_ChildItemPopup
		ggoSpread.SSSetEdit 	C_ChildItemNm,	"��üǰ���", 30
		ggoSpread.SSSetEdit 	C_ChildSpec,	"�԰�"		, 30
		ggoSpread.SSSetEdit 	C_ChildAcctNm,	"ǰ�����"	, 10
		ggoSpread.SSSetEdit 	C_ChildProcurTypeNm,"���ޱ���", 10
		ggoSpread.SSSetFloat	C_ChildItemQty,	"��ǰ����ؼ�", 15, "8" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_ChildUnit,	"����"		, 6,,,3,2
		ggoSpread.SSSetButton 	C_ChildUnitPopup
		ggoSpread.SSSetFloat 	C_PrntItemQty,	"��ǰ����ؼ�", 15, "8" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_PrntUnit,		"����"		, 6,,,3,2
		ggoSpread.SSSetButton 	C_PrntUnitPopup
		ggoSpread.SSSetFloat 	C_SafetyLt,		"����L/T"	, 10, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat 	C_LossRate,		"Loss��"	, 10, "7" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetCombo	C_SupplyType,	"�����󱸺�", 8
		ggoSpread.SSSetCombo	C_SupplyTypeNm,	"�����󱸺�", 10
		ggoSpread.SSSetDate		C_ValidFromDt,	"������"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,	"������"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit 	C_ECNNo,		"���躯���ȣ"	, 18,,,18,2
		ggoSpread.SSSetButton 	C_ECNNoPopup
		ggoSpread.SSSetEdit 	C_ECNDesc,		"���躯�泻��"	, 30,,, 100
		ggoSpread.SSSetEdit 	C_ReasonCd,		"���躯��ٰ�"	, 10,,,2, 2
		ggoSpread.SSSetButton 	C_ReasonCdPopup
		ggoSpread.SSSetEdit 	C_ReasonNm,		"���躯��ٰŸ�"	, 14
		ggoSpread.SSSetEdit 	C_Remark,		"���"		, 30,,, 1000
		
		Call ggoSpread.MakePairsColumn(C_ChildItemCd,	C_ChildItemPopup)
		Call ggoSpread.MakePairsColumn(C_ChildUnit,		C_ChildUnitPopup)
		Call ggoSpread.MakePairsColumn(C_PrntUnit,		C_PrntUnitPopup)
		Call ggoSpread.MakePairsColumn(C_ECNNo,			C_ECNNoPopup)
		
		Call ggoSpread.SSSetColHidden(C_SupplyType, C_SupplyType,	True)
		Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols,		True)
		
		ggoSpread.SSSetSplit2(2)
		
		Call SetSpreadLock 
		
		.ReDraw = true    
    End With
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	Dim i
	ggoSpread.Source = frm1.vspdData

	For i=2 To frm1.vspdData.MaxCols
		ggoSpread.SSSetProtected i, -1, -1
	Next
End Sub

'================================== 2.2.5 SetSpreadLock1() ==================================================
' Function Name : SetSpreadLock1
' Function Desc : This method set color and protect in spread sheet celles When An Specific Row is Selected
'=============================================================================================================
Sub SetSpreadLock1(ByVal Row)

    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock		C_ChildItemCd,		Row, C_ChildItemCd,		Row
    ggoSpread.SpreadLock		C_ChildItemPopup,	Row, C_ChildItemPopup,	Row
 	ggoSpread.SpreadLock		C_ChildItemQty,		Row, C_ChildItemQty,	Row
	ggoSpread.SpreadLock		C_ChildUnit,		Row, C_ChildUnit,		Row
	ggoSpread.SpreadLock		C_ChildUnitPopup,	Row, C_ChildUnitPopup,	Row
	ggoSpread.SpreadLock		C_PrntItemQty,		Row, C_PrntItemQty,		Row
	ggoSpread.SpreadLock		C_PrntUnit,			Row, C_PrntUnit,		Row
	ggoSpread.SpreadLock		C_PrntUnitPopup,	Row, C_PrntUnitPopup,	Row
	ggoSpread.SpreadLock		C_SafetyLt,			Row, C_SafetyLt,		Row
	ggoSpread.SpreadLock		C_LossRate,			Row, C_LossRate,		Row
	ggoSpread.SpreadLock		C_SupplyType,		Row, C_SupplyType,		Row
	ggoSpread.SpreadLock		C_SupplyTypeNm,		Row, C_SupplyTypeNm,	Row
	ggoSpread.SpreadLock		C_ValidFromDt,		Row, C_ValidFromDt,		Row
	ggoSpread.SpreadLock		C_ValidToDt,		Row, C_ValidToDt,		Row
	ggoSpread.SpreadLock		C_ECNNo,			Row, C_ECNNo,			Row
	ggoSpread.SpreadLock		C_ECNNoPopup,		Row, C_ECNNoPopup,		Row
	ggoSpread.SpreadLock		C_ECNDesc,			Row, C_ECNDesc,			Row
	ggoSpread.SpreadLock		C_ReasonCd,			Row, C_ReasonCd,		Row
	ggoSpread.SpreadLock		C_ReasonCdPopup,	Row, C_ReasonCdPopup,	Row
	ggoSpread.SpreadLock		C_Remark,			Row, C_Remark,			Row
	
	ggoSpread.SSSetProtected	.vspdData.MaxCols,	Row, Row        
	.vspdData.ReDraw = True
	
	End With

End Sub

'================================== 2.2.6 SetSpreadUnLock() ==================================================
' Function Name : SetSpreadUnLock
' Function Desc : This method set color and protect in spread sheet celles When A Specific Row is Selected
'=============================================================================================================
Sub SetSpreadUnLock(ByVal Row)

    With frm1

    .vspdData.ReDraw = False
	
	ggoSpread.SpreadUnLock	C_ChildItemCd,		Row, C_ChildItemCd,		Row
    ggoSpread.SpreadUnLock	C_ChildItemPopup,	Row, C_ChildItemPopup,	Row
	ggoSpread.SpreadUnLock	C_ChildItemQty,		Row, C_ChildItemQty,	Row
	ggoSpread.SpreadUnLock	C_ChildUnit,		Row, C_ChildUnit,		Row
	ggoSpread.SpreadUnLock	C_ChildUnitPopup,	Row, C_ChildUnitPopup,	Row
	ggoSpread.SpreadUnLock	C_PrntItemQty,		Row, C_PrntItemQty,		Row
	ggoSpread.SpreadUnLock	C_PrntUnit,			Row, C_PrntUnit,		Row
	ggoSpread.SpreadUnLock	C_PrntUnitPopup,	Row, C_PrntUnitPopup,	Row
	ggoSpread.SpreadUnLock	C_SafetyLt,			Row, C_SafetyLt,		Row
	ggoSpread.SpreadUnLock	C_LossRate,			Row, C_LossRate,		Row
	ggoSpread.SpreadUnLock	C_SupplyType,		Row, C_SupplyType,		Row
	ggoSpread.SpreadUnLock	C_SupplyTypeNm,		Row, C_SupplyTypeNm,	Row
	ggoSpread.SpreadUnLock	C_ValidFromDt,		Row, C_ValidFromDt,		Row
	ggoSpread.SpreadUnLock	C_ValidToDt,		Row, C_ValidToDt,		Row
	If frm1.hBomHistoryFlg.value = "Y" Then
	ggoSpread.SpreadUnLock	C_ECNNo,			Row, C_ECNNo,			Row
	ggoSpread.SpreadUnLock	C_ECNNoPopup,		Row, C_ECNNoPopup,		Row
	ggoSpread.SpreadUnLock	C_ReasonCd,			Row, C_ReasonCd,		Row
	ggoSpread.SpreadUnLock	C_ReasonCdPopup,	Row, C_ReasonCdPopup,	Row
	ggoSpread.SpreadUnLock	C_ECNDesc,			Row, C_ECNDesc,			Row
	End If
	ggoSpread.SpreadUnLock	C_Remark,			Row, C_Remark,			Row
	
	ggoSpread.SSSetRequired C_ChildItemCd,	Row, Row
	ggoSpread.SSSetRequired C_ChildItemQty, Row, Row
	ggoSpread.SSSetRequired C_ChildUnit, 	Row, Row
	ggoSpread.SSSetRequired C_PrntItemQty, 	Row, Row
	ggoSpread.SSSetRequired C_PrntUnit, 	Row, Row
	ggoSpread.SSSetRequired C_SupplyType, 	Row, Row
	ggoSpread.SSSetRequired C_SupplyTypeNm,	Row, Row
	ggoSpread.SSSetRequired C_ValidFromDt, 	Row, Row
	ggoSpread.SSSetRequired C_ValidToDt, 	Row, Row
	If frm1.hBomHistoryFlg.value = "Y" Then
	ggoSpread.SSSetRequired C_ECNNo, 		Row, Row
	ggoSpread.SSSetRequired	C_ECNDesc,		Row, Row
	ggoSpread.SSSetRequired	C_ReasonCd,		Row, Row
	End If
	
	.vspdData.Row = Row
	.vspdData.Col = C_EcnNo
	.vspdData.Text = ""
	.vspdData.Col = C_EcnDesc
	.vspdData.Text = ""
	.vspdData.Col = C_ReasonCd
	.vspdData.Text = ""
	.vspdData.Col = C_ReasonNm
	.vspdData.Text = ""
	
	.vspdData.ReDraw = True
	
	End With

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
    Dim strCbo
    Dim strCboCd
    Dim strCboNm

	'****************************
    'List Minor code(�����󱸺�)
    '****************************
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2201", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_SupplyType
    ggoSpread.SetCombo strCboNm,C_SupplyTypeNm
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'=====================================================================================================
Sub InitComboData(ByVal pStart, ByVal pEnd)
	Dim intRow
	Dim intIndex

	With frm1.vspdData
		If pStart = -1	Then pStart = 1
		If pEnd = -1	Then pEnd = .MaxRows	
	
		For intRow = pStart To pEnd
			.Row = intRow
			.Col = C_SupplyType
			intIndex = .value
			.col = C_SupplyTypeNm
			.value = intindex
		Next	
	End With
End Sub

'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()
	C_Select			= 1
	C_PrntItemCd		= 2
	C_ItemNm			= 3
	C_Spec				= 4
	C_AcctNm			= 5
	C_ProcurTypeNm		= 6
	C_ChildSeq			= 7
	C_ChildItemCd		= 8
	C_ChildItemPopup	= 9
	C_ChildItemNm		= 10
	C_ChildSpec			= 11
	C_ChildAcctNm		= 12
	C_ChildProcurTypeNm	= 13
	C_ChildItemQty		= 14
	C_ChildUnit			= 15
	C_ChildUnitPopup	= 16
	C_PrntItemQty		= 17
	C_PrntUnit			= 18
	C_PrntUnitPopup		= 19
	C_SafetyLt			= 20
	C_LossRate			= 21
	C_SupplyType		= 22
	C_SupplyTypeNm		= 23
	C_ValidFromDt		= 24
	C_ValidToDt			= 25
	C_ECNNo				= 26
	C_ECNNoPopup		= 27
	C_ECNDesc			= 28
	C_ReasonCd			= 29
	C_ReasonCdPopup		= 30
	C_ReasonNm			= 31
	C_Remark			= 32
End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)

    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_Select			= iCurColumnPos(1)
			C_PrntItemCd		= iCurColumnPos(2)
			C_ItemNm			= iCurColumnPos(3)
			C_Spec				= iCurColumnPos(4)
			C_AcctNm			= iCurColumnPos(5)
			C_ProcurTypeNm		= iCurColumnPos(6)
			C_ChildSeq			= iCurColumnPos(7)
			C_ChildItemCd		= iCurColumnPos(8)
			C_ChildItemPopup	= iCurColumnPos(9)
			C_ChildItemNm		= iCurColumnPos(10)
			C_ChildSpec			= iCurColumnPos(11)
			C_ChildAcctNm		= iCurColumnPos(12)
			C_ChildProcurTypeNm	= iCurColumnPos(13)
			C_ChildItemQty		= iCurColumnPos(14)
			C_ChildUnit			= iCurColumnPos(15)
			C_ChildUnitPopup	= iCurColumnPos(16)
			C_PrntItemQty		= iCurColumnPos(17)
			C_PrntUnit			= iCurColumnPos(18)
			C_PrntUnitPopup		= iCurColumnPos(19)
			C_SafetyLt			= iCurColumnPos(20)
			C_LossRate			= iCurColumnPos(21)
			C_SupplyType		= iCurColumnPos(22)
			C_SupplyTypeNm		= iCurColumnPos(23)
			C_ValidFromDt		= iCurColumnPos(24)
			C_ValidToDt			= iCurColumnPos(25)
			C_ECNNo				= iCurColumnPos(26)
			C_ECNNoPopup		= iCurColumnPos(27)
			C_ECNDesc			= iCurColumnPos(28)
			C_ReasonCd			= iCurColumnPos(29)
			C_ReasonCdPopup		= iCurColumnPos(30)
			C_ReasonNm			= iCurColumnPos(31)
			C_Remark			= iCurColumnPos(32)
	
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

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function
    
	IsOpenPop = True

	arrParam(0) = "�����˾�"					' �˾� ��Ī 
	arrParam(1) = "B_PLANT"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"						' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"						' Field��(0)
    arrField(1) = "PLANT_NM"						' Field��(1)
    
    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtPlantCd.Focus
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then Exit Function

	If UCase(Trim(frm1.txtPlantCd.value)) = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "ǰ��", "X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	'---------------------------------------------
	' Parameter Setting
	'--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBomType.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	arrParam(5) = "BOM Type"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "BOM Type"					' Header��(0)
    arrHeader(1) = "BOM Ư��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtBomType.Focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(6), arrField(10)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)							' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)							' Item Code
	arrParam(2) = ""												' Combo Set Data:"1029!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""													' Default Value
		
	arrField(0) = 1		'ITEM_CD
    arrField(1) = 2 	'ITEM_NM											
    arrField(2) = 5		'ITEM_ACCT
    arrField(3) = 9 	'PROC_TYPE
    arrField(4) = 4 	'BASIC_UNIT
    arrField(5) = 51	'SINGLE_ROUT_FLG
    arrField(6) = 52	'Major_Work_Center
    arrField(7) = 13	'Phantom_flg
    arrField(8) = 18	'valid_from_dt
    arrField(9) = 19	'valid_to_dt
    arrField(10) = 3	' Field��(1) : "SPECIFICATION"
  
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtItemCd.Focus
		
End Function

'------------------------------------------  OpenItemCd1()  -------------------------------------------------
'	Name : OpenItemCd1()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd1(ByVal pItemCd)

	Dim arrRet
	Dim arrParam(6), arrField(10)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.hPlantCd.value)	' Plant Code
	arrParam(1) = Trim(pItemCd)				' Item Code
	arrParam(2) = ""						' Combo Set Data:"1029!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""						' Default Value
		
    arrField(0) = 1							'ITEM_CD
    arrField(1) = 2 						'ITEM_NM											
    arrField(2) = 6							'ITEM_ACCT
    arrField(3) = 10 						'PROC_TYPE
    arrField(4) = 4 						'BASIC_UNIT
    arrField(5) = 51						'SINGLE_ROUT_FLG
    arrField(6) = 52						'Major_Work_Center
    arrField(7) = 13						'Phantom_flg
    arrField(8) = 18						'valid_from_dt
    arrField(9) = 19						'valid_to_dt
    arrField(10) = 3						'Spec
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd1(arrRet)
	End If
	
	Call SetActiveCell(frm1.vspdData,C_ChildItemCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function
'------------------------------------------  OpenECNInfo()  ----------------------------------------------
'	Name : OpenECNInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenECNInfo(ByVal pEcnNo)

	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(pEcnNo)				' ECNNo
	arrParam(1) = ""						' ReasonCd
	arrParam(2) = ""						' Status
	arrParam(3) = ""						' EBomFlg
	arrParam(4) = ""						' MBomFlg

	iCalledAspName = AskPRAspName("P1410PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetECNInfo(arrRet)
	End If
	
	Call SetActiveCell(frm1.vspdData,C_EcnNo,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenMassRepRef()  ----------------------------------------------
'	Name : OpenMassRepRef()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenMassRepRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	Dim pvRow,pvCol

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("P1413RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1413RA1", "X")
		IsOpenPop = False
		Msgbox "ERROR: Not Found The Asp Name."
		Exit Function
	End If
	   
	arrParam(0) = frm1.hPlantCd.value	'Plant Cd
	arrParam(1) = ""					'Child Item Cd - ��üǰ 
	arrParam(2) = frm1.hBomType.value	'Bom Type
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=330px; center: Yes; help: No; resizable: No; status: No;")
	
	pvRow = frm1.vspdData.ActiveRow
	pvCol = frm1.vspdData.ActiveCol
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetMassRep(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,pvCol,pvRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function
'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit(ByVal pUnit, ByVal pCol)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(pUnit)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet,pCol)
	End If	
	
	Call SetActiveCell(frm1.vspdData,pCol,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenReasonPopup()  ------------------------------------------
'	Name : OpenReasonPopup()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonPopup(ByVal pReasonCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
  
	'---------------------------------------------
	' Parameter Setting
	'--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "���躯���ȣ�˾�"					' �˾� ��Ī 
	arrParam(1) = "B_MINOR"								' TABLE ��Ī 
	arrParam(2) = UCase(Trim(pReasonCd))				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1402", "''", "S") & ""
	
	arrParam(5) = "���躯��ٰ�"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "���躯��ٰ�"					' Header��(0)
    arrHeader(1) = "���躯��ٰŸ�"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReasonInfo(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_ReasonCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.value = UCase(Trim(arrRet(0)))
	frm1.txtItemNm.value = Trim(arrRet(1))
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd1(byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ChildItemCd
	frm1.vspdData.Text = UCase(Trim(arrRet(0)))
	frm1.vspdData.Col = C_ChildItemNm
	frm1.vspdData.Text = arrRet(1)
	frm1.vspdData.Col = C_ChildSpec
	frm1.vspdData.Text = arrRet(10)	
	frm1.vspdData.Col = C_ChildAcctNm
	frm1.vspdData.Text = arrRet(2)	
	frm1.vspdData.Col = C_ChildProcurTypeNm
	frm1.vspdData.Text = arrRet(3)	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlantCd(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	frm1.txtBomType.Value	= arrRet(0)
	frm1.txtBomType.focus
	Set gActiveElement = document.activeElement  		
End Function

'------------------------------------------  SetECNInfo()  ------------------------------------------------
'	Name : SetECNInfo()
'	Description : ECNInfo Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetECNInfo(byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_EcnNo
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_EcnDesc
	frm1.vspdData.Text = arrRet(1)
	frm1.vspdData.Col = C_ReasonCd
	frm1.vspdData.Text = arrRet(2)
	frm1.vspdData.Col = C_ReasonNm
	frm1.vspdData.Text = arrRet(3)	
End Function

'------------------------------------------  SetUnit()  ------------------------------------------------
'	Name : SetUnit()
'	Description : Open Unit Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetUnit(ByVal arrRet, ByVal pCol)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = pCol
	frm1.vspdData.Text = arrRet(0)
End Function

'------------------------------------------  SetMassRep(arrRet)  --------------------------------------
'	Name : SetMassRep(arrRet)
'	Description : Open Mess Change Reference���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Sub SetMassRep(arrRet)
	Dim iRow
	With frm1.vspdData
		.ReDraw = False		
		For iRow = 1 To .MaxRows
			Call udf_RowUpdate(iRow)	'UnLock
			.Row = iRow
			.Col = C_Select
			.value = 1
			.Col = C_ChildItemCd
			.value = arrRet(1)
			.Col = C_ChildItemNm
			.value = arrRet(2)
			.Col = C_ChildSpec
			.value = arrRet(5)
			.Col = C_ChildAcctNm
			.value = arrRet(4)
			.Col = C_ChildProcurTypeNm
			.value = arrRet(6)
			.Col = C_ChildItemQty
			.value = arrRet(8)
			.Col = C_ChildUnit
			.value = udf_AlterValue(.text, arrRet(9))		'User Defined Function
			.Col = C_PrntItemQty
			.value = arrRet(10)
			.Col = C_PrntUnit
			.value = udf_AlterValue(.text, arrRet(11))
			.Col = C_SafetyLt
			.value = arrRet(12)
			.Col = C_LossRate
			.value = arrRet(13)
			.Col = C_SupplyType
			.text = udf_AlterValue(.text, arrRet(7))
			.Col = C_ValidFromDt
			.text = udf_AlterValue(.text, arrRet(14))			
			.Col = C_ValidToDt
			.text = udf_AlterValue(.text, arrRet(15))
			.Col = C_ECNNo
			.value = arrRet(16)
			.Col = C_ECNDesc
			.value = arrRet(17)
			.Col = C_ReasonCd
			.value = arrRet(18)
			.Col = C_ReasonNm
			.value = arrRet(19)
			.Col = C_Remark
			.value = udf_AlterValue(.text, arrRet(20))
			Call InitComboData(iRow, iRow)
		Next
		.ReDraw = True
	End With
	lgBlnFlgChgValue = True
End Sub
'------------------------------------------  SetReasonInfo()  --------------------------------------------------
'	Name : SetReasonInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function SetReasonInfo(byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ReasonCd
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_ReasonNm
	frm1.vspdData.Text = arrRet(1)
	
	lgBlnFlgChgValue = True
End Function

'==========================================================================================
'   Event Name : LookUpEcnInfo
'   Event Desc : EcnNo Change Event�߻��� ��ȸ 
'==========================================================================================
Sub LookUpEcnInfo(ByVal pEcnNo,ByVal pReasonCd,ByVal pTarget)
	Dim strVal

    Err.Clear                                                               <%'��: Protect system from crashing%>

	If Trim(pEcnNo) = "" AND Trim(pTarget) = "ALL" Then Exit Sub
	If Trim(pReasonCd) = "" AND Trim(pTarget) = "REASON" Then Exit Sub
	
	If   LayerShowHide(1) = False Then Exit Sub

	strVal = BIZ_PGM_LOOKUP_ECN_INFO & "?txtMode=" & parent.UID_M0001								<%'��: �����Ͻ� ó�� ASP�� ���� %>
	strVal = strVal & "&txtEcnNo=" & Trim(pEcnNo)
	strVal = strVal & "&txtReasonCd=" & Trim(pReasonCd)
	strVal = strVal & "&txtTarget=" & Trim(pTarget)
	strVal = strVal & "&Row=" & frm1.vspdData.Row

	Call RunMyBizASP(MyBizASP, strVal)
End Sub

Function LookUpEcnInfoOk(ByVal pReasonCd,ByVal pReasonNm,ByVal pEcnDesc,ByVal pTarget,ByVal Row)
	With frm1.vspdData
		.ReDraw = False
		.Row = CLng(Row)
		
		If pTarget = "ALL" Then
			.Col = C_ReasonCd
			.Text = Trim(pReasonCd)
			.Col = C_ReasonNm
			.Text = Trim(pReasonNm)
			.Col = C_EcnDesc
			.Text = pEcnDesc
		ElseIf pTarget = "REASON" Then
			.Col = C_ReasonNm
			.Text = Trim(pReasonNm)
		End If
			
		.ReDraw = True
	End With
End Function

'==========================================================================================
'   Event Name : LookUpItemInfo
'   Event Desc : ChildItemCd Change Event�߻��� ��ȸ 
'==========================================================================================
Sub LookUpItemInfo(ByVal pItemCd)
	Dim strVal

    Err.Clear                                                               <%'��: Protect system from crashing%>

	If Trim(pItemCd) = "" Then Exit Sub
	If   LayerShowHide(1) = False Then Exit Sub

	strVal = BIZ_PGM_LOOKUP_ITEM_INFO & "?txtMode=" & parent.UID_M0001								<%'��: �����Ͻ� ó�� ASP�� ���� %>
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			<%'��: ��ȸ ���� ����Ÿ %>
	strVal = strVal & "&txtItemCd=" & Trim(pItemCd)							<%'��: ��ȸ ���� ����Ÿ %>
	strVal = strVal & "&Row=" & frm1.vspdData.Row							'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
End Sub

Function LookUpItemInfoOk(ByVal pItemNm, ByVal pSpec, ByVal pAcctNm, ByVal pProcurTypeNm, ByVal Row)
	With frm1.vspdData
		.ReDraw = False
		.Row = CLng(Row)
		
		.Col = C_ChildItemNm
		.Text = pItemNm
		.Col = C_ChildSpec
		.Text = pSpec
		.Col = C_ChildAcctNm
		.Text = pAcctNm
		.Col = C_ChildProcurTypeNm
		.Text = pProcurTypeNm
			
		.ReDraw = True
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
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet

       '----------  Coding part  -------------------------------------------------------------
    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call InitComboBox
    Call SetToolBar("1100000000011")										'��: ��ư ���� ���� 

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
'   Event Name : txtFromReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFromReqDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtFromReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtToReqDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtToReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtDoDate_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDoDate_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDoDate.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtDoDate.Focus 
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtFromReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromReqDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToReqDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then							'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
    Else

    End If
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

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
	With frm1.vspdData
		.Row = Row
		Select Case Col
			Case C_EcnNo
				.Col = Col
				Call LookUpEcnInfo(.text,"", "ALL")
			Case  C_ReasonCd
				.Col = Col
				Call LookUpEcnInfo("",.text, "REASON")
			Case C_ChildItemCd
				.Col = Col
				Call LookUpItemInfo(.text)
		End Select
    End With

	lgBlnFlgChgValue = True    
End Sub

'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)

End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row
		Select Case Col
			Case  C_SupplyTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_SupplyType
				.Value = intIndex
		End Select
    End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData
		.Row = Row
		ggoSpread.Source = frm1.vspdData

		If Col = C_Select Then
			If ButtonDown = 1 Then
				Call udf_RowUpdate(Row)		'User Defined Function
			Else
				Call udf_RowUnDo(Row)		'User Defined Function
			End If
		End If

		Select Case Col
			Case C_ChildItemPopup
				.Col = C_ChildItemCd
				.Row = Row
				Call OpenItemCd1(.Text)
				Call SetActiveCell(frm1.vspdData,C_ChildItemCd,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
			Case C_ChildUnitPopup
				.Col = C_ChildUnit
				.Row = Row
				Call OpenUnit(.Text,C_ChildUnit)
				Call SetActiveCell(frm1.vspdData,C_ChildUnit,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
			Case C_PrntUnitPopup
				.Col = C_PrntUnit
				.Row = Row
				Call OpenUnit(.Text,C_PrntUnit)
				Call SetActiveCell(frm1.vspdData,C_PrntUnit,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
			Case C_ECNNoPopup
				.Col = C_ECNNo
				.Row = Row
				Call OpenECNInfo(.Text)
				Call SetActiveCell(frm1.vspdData,C_ECNNo,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
			Case C_ReasonCdPopup
				.Col = C_ReasonCd
				.Row = Row
				Call OpenReasonPopup(.Text)
				Call SetActiveCell(frm1.vspdData,C_ReasonCd,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
		End Select
	End With
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

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	If Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantNm.value = "" 
	End If	
	
	If Trim(frm1.txtItemCd.value) = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
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
	Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    On Error Resume Next                                                    '��: Protect system from crashing
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '��: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
	'If Not chkField(Document, "2") Then
    '   Exit Function
    'End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Precheck area
    '-----------------------

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     							                                      '��: Save db data
    
    FncSave = True 
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
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
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
			strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.hPlantCd.value))			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.hItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtBomType="		& UCase(Trim(.hBomType.valud))
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.txtPlantCd.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.txtItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtBomType="		& UCase(Trim(.txtBomType.value))		'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
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
	Call InitComboData(-1,-1)
	Call SetToolBar("11001000000111")														'��: ��ư ���� ���� 
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim IntRows 
    Dim strVal
	Dim lGrpCnt
	Dim iColSep
	Dim TmpBuffer
	Dim iTotalStr
	
	On Error Resume Next
	Err.Clear
	
    DbSave = False                                                          '��: Processing is NG
	
    LayerShowHide(1)
	
	iColSep = Parent.gColSep
	lGrpCnt = 1
	ReDim TmpBuffer(0)
	
    With frm1
		.txtMode.Value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.Value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
		
    '-----------------------
    'Data manipulate area
    '-----------------------
    For IntRows = 1 To .vspdData.MaxRows
		.vspdData.Row = IntRows
		.vspdData.Col = 0
		Select Case .vspdData.Text
            Case ggoSpread.UpdateFlag
				strVal = ""
				strVal = strVal & "U" & iColSep	& IntRows & iColSep			'��: U=Update, RowNum
                .vspdData.Col = C_PrntItemCd	'2
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ChildItemCd	'3
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ChildSeq		'4
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ChildItemQty	'5
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ChildUnit		'6
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_PrntItemQty	'7
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_PrntUnit		'8
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_SafetyLt		'9
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_LossRate		'10
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_SupplyType	'11
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ValidFromDt	'12
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ValidToDt		'13
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ECNNo			'14
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ECNDesc		'15
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_ReasonCd		'16
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_Remark		'17
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                
                ReDim Preserve TmpBuffer(lGrpCnt-1)
                TmpBuffer(lGrpCnt-1) = strVal
                lGrpCnt = lGrpCnt + 1
                
        End Select
    Next
	iTotalStr = Join(TmpBuffer, "")
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = iTotalStr

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: ���� �����Ͻ� ASP �� ���� 
    DbSave = True                                                           '��: Processing is OK
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	frm1.txtPlantCd.value = frm1.hPlantCd.value
	frm1.txtItemCd.value = frm1.hItemCd.value
	frm1.txtBomType.value = frm1.hBomType.value
	
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
	IsOpenPop = False
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
    Call InitSpreadSheet
    ggoSpread.Source = gActiveSpdSheet
		
	Call ggoSpread.ReOrderingSpreadData()
End Sub 

'########################################################################################
' User Defined Functions
'########################################################################################
'========================================================================================
' Function Name : udf_RowUpdate(ByVal Row)
' Function Desc : Row�� Update�Ѵ�.
'========================================================================================
Sub udf_RowUpdate(ByVal pRow)
	ggoSpread.UpdateRow pRow
	Call SetSpreadUnLock(pRow)
End Sub

'========================================================================================
' Function Name : udf_RowUnDo(ByVal Row)
' Function Desc : Row�� ���(�������)�Ѵ�.
'========================================================================================
Sub udf_RowUnDo(ByVal pRow)
	Call ggoSpread.EditUndo(pRow, pRow)
	Call InitComboData(pRow, pRow)
	ggoSpread.SSDeleteFlag pRow,pRow
	Call SetSpreadLock1(pRow)
End Sub

'========================================================================================
' Function Name : udf_AlterValue(ByVal pOld, ByVal pChkVal)
' Function Desc : pChkVal�� ���� ������ ������(pOld)�� �����Ѵ�.
'========================================================================================
Function udf_AlterValue(ByVal pOld, ByVal pChkVal)
	If Trim(pChkVal) = "" Then
		udf_AlterValue = pOld
	Else
		udf_AlterValue = pChkVal
	End If
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ���ϰ���ü</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenMassRepRef()">�ϰ���ü</A></TD>
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
							<TABLE WIDTH=100% CELLSPACING=0>					
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD></TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="��ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="��ǰ��"></TD>
									<TD CLASS=TD5 NOWRAP>BOM Type</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBomType" SIZE=5 MAXLENGTH=3 tag="12xxxU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBomNo"></TD>
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
							<TR>
								<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcct" SIZE=25 MAXLENGTH=3 tag="24xxxU" ALT="ǰ�����"></TD>
								<TD CLASS=TD5 NOWRAP>���ش���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBaseUnit" SIZE=5 MAXLENGTH=4 tag="24xxxU" ALT="���ش���"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�԰�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpec" SIZE=40 MAXLENGTH=50 tag="24xxxU" ALT="�԰�"></TD>
								<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="14" ALT="������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="14" ALT="������"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="22" TITLE="SPREAD" id=OBJECT4> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBomType" tag="24">
<INPUT TYPE=HIDDEN NAME="hBomHistoryFlg" tag="24" value="Y">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
