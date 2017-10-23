<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1601ma1.asp
'*  4. Program Name         : Copy Item by Plant
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Ryu Sung Won
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance
On Error Resume Next
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY_ID	= "p1601mb1.asp"												'��: Detail Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID	= "p1601mb2.asp"												'��: Detail Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_LOOKUPVATTYPE_ID	= "p1601mb3.asp"
'==========================================================================================================
'==========================================================================================================

Dim C_Select
Dim C_Item
Dim C_ItmNm
Dim C_PrcCtrlInd
Dim C_PrcCtrlIndNm	
Dim C_UnitPrice		
Dim C_IBPValidFromDt	
Dim C_IBPValidToDt	
Dim C_ItmFormalNm
Dim C_ItmAcc
Dim C_Unit
Dim C_UnitPopup
Dim C_ItmGroupCd
Dim C_ItmGroupPopup
Dim C_ItmGroupNm
Dim C_Phantom
Dim C_BlanketPur
Dim C_BaseItm
Dim C_BaseItmPopup
Dim C_BaseItmNm
Dim C_SumItmClass
Dim C_DefaultFlg
Dim C_PicFlg
Dim C_ItmSpec
Dim C_UnitWeight
Dim C_UnitOfWeight
Dim C_WeightUnitPopup
Dim C_UnitGrossWeight
Dim C_UnitOfGrossWeight
Dim C_GrossUnitPopup
Dim C_CBM
Dim C_CBMDesc
Dim C_DrawNo
Dim C_HsCd
Dim C_HsCdPopup
Dim C_HsUnit
Dim C_VatType
Dim C_VatTypePopup
Dim C_VatRate
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_HdnSumItmClass
Dim C_HdnItmAcc

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgInsrtFlg
Dim lgFlgAllSelected		'When Selected All
Dim lgFlgCancelClicked		'Cancel Button Clicked
Dim lgFlgCopyClicked		'Copy Button Clicked
Dim lgFlgBtnSelectAllClicked 'When btnSelectAll Clicked

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim ihGridCnt                     'hidden Grid Row Count
Dim intItemCnt                    'hidden Grid Row Count
Dim IsOpenPop					 'Popup
Dim gSelframeFlg
Dim lgRdoOldVal
Dim iDBSYSDate
Dim StartDate, EndDate
'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_Select			= 1
	C_Item				= 2
	C_ItmNm				= 3
	C_ItmSpec			= 4
	C_PrcCtrlInd		= 5
	C_PrcCtrlIndNm		= 6
	C_UnitPrice			= 7
	C_IBPValidFromDt	= 8	
	C_IBPValidToDt		= 9
	C_ItmFormalNm		= 10
	C_ItmAcc			= 11
	C_Unit				= 12
	C_UnitPopup			= 13
	C_ItmGroupCd		= 14
	C_ItmGroupPopup		= 15 
	C_ItmGroupNm		= 16
	C_Phantom			= 17
	C_BlanketPur		= 18 
	C_BaseItm			= 19
	C_BaseItmPopup		= 20
	C_BaseItmNm			= 21
	C_SumItmClass		= 22
	C_DefaultFlg		= 23
	C_PicFlg			= 24
	C_UnitWeight		= 25 
	C_UnitOfWeight		= 26 
	C_WeightUnitPopup	= 27
	C_UnitGrossWeight	= 28 
	C_UnitOfGrossWeight	= 29
	C_GrossUnitPopup	= 30
	C_CBM				= 31 
	C_CBMDesc			= 32
	C_DrawNo			= 33
	C_HsCd				= 34
	C_HsCdPopup			= 35
	C_HsUnit			= 36
	C_VatType			= 37
	C_VatTypePopup		= 38
	C_VatRate			= 39
	C_ValidFromDt		= 40
	C_ValidToDt			= 41
	C_HdnSumItmClass	= 42
	C_HdnItmAcc			= 43
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count

	'frm1.btnCopy.disabled = True
	'frm1.btnSelectAll.disabled = True
	frm1.btnSelectAll.value = "��ü����"
	lgFlgAllSelected = False
	lgFlgCancelClicked = False
	lgFlgCopyClicked = False
	lgFlgBtnSelectAllClicked = False

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

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
    With frm1.vspdData
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030601",,parent.gAllowDragDropSpread   
	
    .MaxCols = C_HdnItmAcc + 1
    .MaxRows = 0
    
	.ReDraw = false
	 
	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetCheck	C_Select ,		"",				2,,,1
	ggoSpread.SSSetEdit 	C_Item,			"ǰ��",		15,,,18,2
	ggoSpread.SSSetEdit 	C_ItmNm,		"ǰ���",	25,,,40
	ggoSpread.SSSetEdit 	C_ItmSpec,		"�԰�",		25,,,40
	ggoSpread.SSSetEdit 	C_ItmFormalNm,	"ǰ�����ĸ�",25,,,60
	ggoSpread.SSSetCombo 	C_PrcCtrlInd,	"�ܰ�����", 12
	ggoSpread.SSSetCombo 	C_PrcCtrlIndNm, "�ܰ�����", 12
	ggoSpread.SSSetFloat	C_UnitPrice,	"�ܰ�",		16,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetCombo 	C_ItmAcc,		"ǰ�����", 12
	ggoSpread.SSSetEdit 	C_Unit,			"����",		10,,,3,2
	ggoSpread.SSSetButton 	C_UnitPopup
	ggoSpread.SSSetEdit 	C_ItmGroupCd,	"ǰ��׷�",	10,,,10,2
	ggoSpread.SSSetButton 	C_ItmGroupPopup
	ggoSpread.SSSetEdit 	C_ItmGroupNm,	"ǰ��׷��",16
	ggoSpread.SSSetCombo 	C_Phantom,		"����",		10,2
	ggoSpread.SSSetCombo 	C_BlanketPur,	"���ձ���", 10,2
	ggoSpread.SSSetEdit 	C_BaseItm,		"����ǰ��",	15,,,18,2
	ggoSpread.SSSetButton 	C_BaseItmPopup
	ggoSpread.SSSetEdit 	C_BaseItmNm,	"����ǰ���",15
	ggoSpread.SSSetCombo 	C_SumItmClass,	"ǰ��Ŭ����",15
	ggoSpread.SSSetCombo 	C_DefaultFlg,	"��ȿ����",	10,2
	ggoSpread.SSSetEdit 	C_PicFlg,		"��������",	10,2
	ggoSpread.SSSetFloat	C_UnitWeight,	"Net�߷�",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit 	C_UnitOfWeight, "Net����",	10,,,3,2
	ggoSpread.SSSetButton 	C_WeightUnitPopup
	ggoSpread.SSSetFloat	C_UnitGrossWeight,	 "Gross�߷�",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit 	C_UnitOfGrossWeight, "Gross����",	10,,,3,2
	ggoSpread.SSSetButton 	C_GrossUnitPopup
	ggoSpread.SSSetFloat	C_CBM,			"CBM(����)",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit 	C_CBMDesc,		"CBM����",	15,,,50	
	ggoSpread.SSSetEdit 	C_DrawNo,		"�����ȣ",	15,,,20
	ggoSpread.SSSetEdit 	C_HsCd,			"HS�ڵ�",	15,,,20,2
	ggoSpread.SSSetButton 	C_HsCdPopup
	ggoSpread.SSSetEdit 	C_HsUnit,		"HS����",	10,,,3,2
	ggoSpread.SSSetEdit 	C_VatType,		"VAT����",	10,,,3,2
	ggoSpread.SSSetButton 	C_VatTypePopup
	ggoSpread.SSSetFloat	C_VatRate,		"VAT��",	12, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetDate		C_IBPValidFromDt,"��ȿ������",	12, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_IBPValidToDt,	"��ȿ������",	12, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_ValidFromDt,	"ǰ�������",	12, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_ValidToDt,	"ǰ��������",	12, 2, parent.gDateFormat
	ggoSpread.SSSetCombo 	C_HdnSumItmClass,"�����ǰ��Ŭ����",15
	ggoSpread.SSSetCombo 	C_HdnItmAcc,	"ǰ�����",		15
	
	call ggoSpread.MakePairsColumn(C_BaseItm,		C_BaseItmPopup)
	call ggoSpread.MakePairsColumn(C_ItmGroupCd,	C_ItmGroupPopup)
	call ggoSpread.MakePairsColumn(C_Unit,			C_UnitPopup)
	call ggoSpread.MakePairsColumn(C_UnitOfWeight,	C_WeightUnitPopup)
	call ggoSpread.MakePairsColumn(C_UnitGrossWeight,	C_GrossUnitPopup)
	call ggoSpread.MakePairsColumn(C_HsCd,			C_HsCdPopup)
	call ggoSpread.MakePairsColumn(C_VatType,		C_VatTypePopup)
	
	Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols,		True)
	Call ggoSpread.SSSetColHidden(.MaxCols - 1,	.MaxCols - 1,	True)
	Call ggoSpread.SSSetColHidden(.MaxCols - 2,	.MaxCols - 2,	True)
	Call ggoSpread.SSSetColHidden(C_BaseItmNm,	C_BaseItmNm,	True)
	Call ggoSpread.SSSetColHidden(C_PrcCtrlInd,	C_PrcCtrlInd,	True)
	
	ggoSpread.SSSetSplit2(2)										'frozen ����߰� 
	Call SetSpreadLock 
    
    .ReDraw = true
    
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	Dim i
	
	For i=2 To C_HdnItmAcc
		ggoSpread.SSSetProtected i, -1, -1
	Next
	ggoSpread.SSSetProtected frm1.vspdData.MaxCols, -1
End Sub


'================================== 2.2.6 SetSpreadUnLock() ==================================================
' Function Name : SetSpreadUnLock
' Function Desc : This method set color and protect in spread sheet celles When A Specific Row is Selected
'=============================================================================================================

Sub SetSpreadUnLock(ByVal Col, ByVal Row)

	ggoSpread.SpreadUnLock	C_PrcCtrlIndNm,		Row, C_PrcCtrlIndNm,	Row
	ggoSpread.SpreadUnLock	C_UnitPrice,		Row, C_UnitPrice,		Row
	ggoSpread.SpreadUnLock	C_IBPValidFromDt,	Row, C_IBPValidFromDt,	Row
	ggoSpread.SpreadUnLock	C_IBPValidToDt,		Row, C_IBPValidToDt,	Row
    
	ggoSpread.SSSetRequired 	C_PrcCtrlIndNm, 	Row, Row
	ggoSpread.SSSetRequired 	C_UnitPrice, 		Row, Row
	ggoSpread.SSSetRequired		C_IBPValidFromDt,	Row, Row
	ggoSpread.SSSetRequired		C_IBPValidToDt,		Row, Row

End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetProtected	C_Select,		pvStartRow, pvEndRow

	ggoSpread.SSSetRequired 	C_Item, 		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_ItmNm, 		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_PrcCtrlIndNm,	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_UnitPrice, 	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_ItmAcc, 		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_Unit,			pvStartRow, pvEndRow
	    
	ggoSpread.SSSetProtected	C_ItmGroupNm,	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_Phantom,		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_BlanketPur,	pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_BaseItmNm,	pvStartRow, pvEndRow
	ggoSpread.SSSetRequired		C_DefaultFlg,	pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_PicFlg,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_HsUnit,		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_IBPValidFromDt,pvStartRow, pvEndRow
	ggoSpread.SSSetRequired 	C_IBPValidToDt, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_ValidFromDt, 	pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_ValidToDt, 	pvStartRow, pvEndRow
	ggoSpread.SSSetProtected	C_VatRate, 		pvStartRow, pvEndRow		'�߰��� �κ� 2002-10-11
	ggoSpread.SSSetProtected	frm1.vspdData.MaxCols, pvStartRow, pvEndRow

End Sub

'================================== 2.2.5 SetSpreadLock1() ==================================================
' Function Name : SetSpreadLock1
' Function Desc : This method set color and protect in spread sheet celles When An Specific Row is Selected
'=============================================================================================================

Sub SetSpreadLock1(ByVal Col, ByVal Row)

 	ggoSpread.SpreadLock		C_PrcCtrlIndNm,		Row, C_PrcCtrlIndNm,	Row
	ggoSpread.SpreadLock		C_UnitPrice,		Row, C_UnitPrice,		Row
	ggoSpread.SpreadLock		C_IBPValidFromDt,	Row, C_IBPValidFromDt,	Row
	ggoSpread.SpreadLock		C_IBPValidToDt,		Row, C_IBPValidToDt,	Row
	
	ggoSpread.SSSetProtected	C_PrcCtrlIndNm, 	Row, Row
	ggoSpread.SSSetProtected	C_UnitPrice, 		Row, Row
	ggoSpread.SSSetProtected	C_IBPValidFromDt,	Row, Row
	ggoSpread.SSSetProtected	C_IBPValidToDt,		Row, Row

	ggoSpread.SSSetProtected	frm1.vspdData.MaxCols,	Row, Row        
	
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    Dim iColSep
    
    iColSep = parent.gColSep
    
    iPosArr = Split(iPosArr,iColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Select			= iCurColumnPos(1)
			C_Item				= iCurColumnPos(2)
			C_ItmNm				= iCurColumnPos(3)
			C_ItmSpec			= iCurColumnPos(4)
			C_PrcCtrlInd		= iCurColumnPos(5)
			C_PrcCtrlIndNm		= iCurColumnPos(6)
			C_UnitPrice			= iCurColumnPos(7)
			C_IBPValidFromDt	= iCurColumnPos(8)	
			C_IBPValidToDt		= iCurColumnPos(9)
			C_ItmFormalNm		= iCurColumnPos(10)
			C_ItmAcc			= iCurColumnPos(11)
			C_Unit				= iCurColumnPos(12)
			C_UnitPopup			= iCurColumnPos(13)
			C_ItmGroupCd		= iCurColumnPos(14)
			C_ItmGroupPopup		= iCurColumnPos(15)
			C_ItmGroupNm		= iCurColumnPos(16)
			C_Phantom			= iCurColumnPos(17)
			C_BlanketPur		= iCurColumnPos(18)
			C_BaseItm			= iCurColumnPos(19)
			C_BaseItmPopup		= iCurColumnPos(20)
			C_BaseItmNm			= iCurColumnPos(21)
			C_SumItmClass		= iCurColumnPos(22)
			C_DefaultFlg		= iCurColumnPos(23)
			C_PicFlg			= iCurColumnPos(24)
			C_UnitWeight		= iCurColumnPos(25)
			C_UnitOfWeight		= iCurColumnPos(26)
			C_WeightUnitPopup	= iCurColumnPos(27)
			C_UnitGrossWeight	= iCurColumnPos(28) 
			C_UnitOfGrossWeight	= iCurColumnPos(29)
			C_GrossUnitPopup	= iCurColumnPos(30)
			C_CBM				= iCurColumnPos(31) 
			C_CBMDesc			= iCurColumnPos(32)
			C_DrawNo			= iCurColumnPos(33)
			C_HsCd				= iCurColumnPos(34)
			C_HsCdPopup			= iCurColumnPos(35)
			C_HsUnit			= iCurColumnPos(36)
			C_VatType			= iCurColumnPos(37)
			C_VatTypePopup		= iCurColumnPos(38)
			C_VatRate			= iCurColumnPos(39)
			C_ValidFromDt		= iCurColumnPos(40)
			C_ValidToDt			= iCurColumnPos(41)
			C_HdnSumItmClass	= iCurColumnPos(42)
			C_HdnItmAcc			= iCurColumnPos(43)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm

	'****************************
    ' ����� ǰ��Ŭ���� 
    '****************************
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1002", "''", "S") & " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	Call SetCombo2(frm1.cboItemClass, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnSumItmClass
    ggoSpread.SetCombo strCboNm,C_SumItmClass
    
    '****************************
    ' ǰ����� 
    '****************************     
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnItmAcc
    ggoSpread.SetCombo strCboNm,C_ItmAcc

	'****************************
    ' ����,���ձ���,��ȿ���� 
    '****************************     
    strCboCd = ""
    strCboNm = ""
	
	strCboCd = "Y" & vbTab & "N"
	
	ggoSpread.SetCombo strCboCd,C_Phantom		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboCd,C_BlanketPur	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboCd,C_DefaultFlg	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	
	'****************************
    'Price Control Ind
    '****************************
	strCboCd = "" 
	strCboNm = ""
	
	ggoSpread.Source = frm1.vspdData

    strCboCd = strCboCd & "S" & vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "ǥ�شܰ�" & vbTab    'Setting Job Nm in Detail Sheet

    strCboCd = strCboCd & "M" '& vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "�̵���մܰ�" '& vbTab            'Setting Job Nm in Detail Sheet

    ggoSpread.SetCombo strCboCd,C_PrcCtrlInd		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboNm,C_PrcCtrlIndNm	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
End Sub

'==========================================  2.2.6 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display in Spread(s)
'========================================================================================================= 
Sub InitSpreadComboBox()
    Dim strCboCd 
    Dim strCboNm
    
    '****************************
    ' ����� ǰ��Ŭ���� 
    '****************************
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1002", "''", "S") & " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	'Call SetCombo2(frm1.cboItemClass, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnSumItmClass
    ggoSpread.SetCombo strCboNm,C_SumItmClass
    
    '****************************
    ' ǰ����� 
    '****************************    
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	'Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnItmAcc
    ggoSpread.SetCombo strCboNm,C_ItmAcc

    '****************************
    ' ����,���ձ���,��ȿ���� 
    '****************************    
    strCboCd = ""
    strCboNm = ""
    
    strCboCd = "Y" & vbTab & "N"
	
	ggoSpread.SetCombo strCboCd,C_Phantom		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboCd,C_BlanketPur	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboCd,C_DefaultFlg	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	
	'****************************
    'Price Control Ind
    '****************************	
	strCboCd = "" 
	strCboNm = ""
	
	ggoSpread.Source = frm1.vspdData

    strCboCd = strCboCd & UCase("S") & vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "ǥ�شܰ�" & vbTab    'Setting Job Nm in Detail Sheet

    strCboCd = strCboCd & UCase("M") & vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "�̵���մܰ�" & vbTab            'Setting Job Nm in Detail Sheet

    ggoSpread.SetCombo strCboCd,C_PrcCtrlInd		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboNm,C_PrcCtrlIndNm	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		'.ReDraw = False
		
		For intRow = lngStartRow To .MaxRows
			If iPos = 1 Then
				.Row = intRow
				.Col = C_HdnItmAcc
				intIndex = .value
				.col = C_ItmAcc
				.value = intindex
				
				.Row = intRow
				.Col = C_HdnSumItmClass
				intIndex = .value
				.col = C_SumItmClass
				.value = intindex
				
				.Row = intRow
				.Col = C_PrcCtrlInd
				intIndex = .value
				.col = C_PrcCtrlIndNm
				.value = intindex
				
			Else
				.Row = intRow
				.Col = C_ItmAcc
				intIndex = .value
				.col = C_HdnItmAcc
				.value = intindex
			
				.Row = intRow
				.Col = C_SumItmClass
				intIndex = .value
				.col = C_HdnSumItmClass
				.value = intindex
				
				.Row = intRow
				.Col = C_PrcCtrlInd
				intIndex = .value
				.col = C_PrcCtrlIndNm
				.value = intindex
			End IF							
		Next	
		'.ReDraw = True
	End With
End Sub

Function SetFieldProp(ByVal lRow, ByVal sType)
	ggoSpread.Source = frm1.vspdData
    
	ggoSpread.SSSetRequired	C_PrcCtrlInd,	lRow, lRow
	ggoSpread.SSSetRequired	C_UnitPrice,	lRow, lRow
End Function

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

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim activateField
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
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
		Call SetConPlant(arrRet, 0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenCondPlant1()  -------------------------------------------------
'	Name : OpenCondPlant1()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim activateField
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd1.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "���ذ���"					' Header��(0)
    arrHeader(1) = "���ذ����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet, 1)
	End If	
	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd1.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(strCode, iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If iPos = 1 Then
		If frm1.txtPlantCd.value = "" Then
			Call DisplayMsgBox("971012","X", "����","x")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If		
	End If
	
	IsOpenPop = True
	
	If iPos = 0 Then
		arrParam(0) = strCode						' Item Code
		arrParam(1) = ""							' Item Name
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
		arrParam(3) = ""							' Default Value
		
		iCalledAspName = AskPRAspName("B1B01PA2")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
			IsOpenPop = False
			Exit Function
		End If
	ElseIf iPos = 1 Then
		arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
		arrParam(1) = strCode						' Item Code
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
		arrParam(3) = ""							' Default Value
		
		iCalledAspName = AskPRAspName("B1B11PA2")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA2", "X")
			IsOpenPop = False
			Exit Function
		End If
	End If

    arrField(0) = 1 								' Field��(0) :"ITEM_CD"
    arrField(1) = 2 								' Field��(1) :"ITEM_NM"
    arrField(2) = 3 								' Field��(2) :"SPEC"
    arrField(3) = 9 								' Field��(2) :"ProcurType"
    arrField(4) = 10 								' Field��(2) :"ProcurType"
        
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet,iPos)
	End If	
	
	Call SetFocusToDocument("M")
	If iPos = 0 Then
		frm1.txtItemCd.focus
	Else
		frm1.txtItemCd1.focus
	End If	
	

End Function

Function OpenUnitPopup(strVal)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = strVal
	arrParam(3) = ""
	arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & " "			
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_Unit,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = frm1.txtHighItemGroupCd.value  
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "ǰ��׷�"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
'    arrField(3) = "LEAF_FLG"	
'    arrField(4) = "UPPER_ITEM_GROUP_CD"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroupCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtHighItemGroupCd.focus
	
End Function

Function OpenItemGroupPopup(strVal)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = strVal
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  AND LEAF_FLG = " & FilterVar("Y", "''", "S") & "  AND VALID_TO_DT >=  " & FilterVar(EndDate , "''", "S") & "" 			
	arrParam(5) = "ǰ��׷�"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
'    arrField(3) = "LEAF_FLG"	
'    arrField(4) = "UPPER_ITEM_GROUP_CD"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_ItmGroupCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

Function OpenBaseItemPopup(strVal)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = strVal						' Item Code
	arrParam(1) = ""							' Item Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = EndDate
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    
    iCalledAspName = AskPRAspName("B1B01PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBasisItemCd(arrRet)
	End If
	
	Call SetActiveCell(frm1.vspdData,C_BaseItm,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

Function OpenWtUnitPopup(strVal)
Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = strVal
	arrParam(3) = ""
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & ""			
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWtUnit(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_UnitOfWeight,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenGrossUnit()  -------------------------------------------
'	Name : OpenGrossUnit()
'	Description : WeightUnit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGrossUnit(byval strWeightUnit)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(strWeightUnit)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & ""			
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetGrossUnit(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_UnitOfGrossWeight,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement

End Function

Function OpenHsPopup(strVal)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HS�˾�"	
	arrParam(1) = "B_HS_CODE"				
	arrParam(2) = strVal
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "HS�ڵ�"
	
    arrField(0) = "HS_CD"	
    arrField(1) = "HS_NM"
    arrField(2) = "HS_SPEC"	
    arrField(3) = "HS_UNIT"
    	
    
    arrHeader(0) = "HS�ڵ�"		
    arrHeader(1) = "HS��"
    arrHeader(2) = "HS�԰�"
    arrHeader(3) = "HS����"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetHSCd(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_HsCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement

End Function

'===========================================================================
' Function Name : OpenBillHdr
' Function Desc : OpenBillHdr Reference Popup
'===========================================================================
Function OpenBillHdr(ByVal VatType)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(1) = "B_MINOR ,B_CONFIGURATION "	' TABLE ��Ī 
	arrParam(2) = VatType				' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & "" _
					& " And b_minor.minor_cd=b_configuration.minor_cd " _
					& " And b_minor.major_cd=b_configuration.major_cd "	_
					& " And b_configuration.SEQ_NO=1 "					' Where Condition
	arrParam(5) = "VAT����"						' TextBox ��Ī 
		
	arrField(0) = "b_minor.MINOR_CD"					' Field��(0)
	arrField(1) = "b_minor.MINOR_NM"					' Field��(1)
	arrField(2) = "F5" & parent.gColSep & "b_configuration.REFERENCE"				' Field��(2)
	    	    
	arrHeader(0) = "VAT����"						' Header��(0)
	arrHeader(1) = "VAT������"					' Header��(1)
	arrHeader(2) = "VAT��"					' Header��(2)

	arrParam(0) = arrParam(5)							' �˾� ��Ī 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBillHdr(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_VatType,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet,ByVal iPos)
	With frm1
		If iPos = 0 Then
			.txtItemCd.value = arrRet(0)
			.txtItemNm.value = arrRet(1)
		Else
			.txtItemCd1.value	= arrRet(0)
			.txtItemNm1.value	= arrRet(1)
			.txtItemSpec1.value = arrRet(2)
			.txtItemProcType1.value = arrRet(4)
			.htxtItemProcType1.value = arrRet(3)
		End If

	End With
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet, byval iPos)
	With frm1
		If iPos = 0 Then
			.txtPlantCd.Value    = arrRet(0)		
			.txtPlantNm.Value    = arrRet(1)
		Else
			.txtPlantCd1.Value    = arrRet(0)		
			.txtPlantNm1.Value    = arrRet(1)
		End If
	End With
End Function
'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetUnit(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_Unit
	frm1.vspdData.Text = arrRet(0)
End Function
'------------------------------------------  SetItemGroup()  --------------------------------------------------
'	Name : SetItemGroup()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetItemGroup(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItmGroupCd
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_ItmGroupNm
	frm1.vspdData.Text = arrRet(1)		
End Function

'------------------------------------------  SetItemGroupCd()  --------------------------------------------------
'	Name : SetItemGroupCd()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetItemGroupCd(Byval arrRet)
	frm1.txtHighItemGroupCd.value = arrRet(0)
	frm1.txtHighItemGroupNm.value = arrRet(0)
End Function

'------------------------------------------  SetBaseItem()  --------------------------------------------------
'	Name : SetBaseItem()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetBasisItemCd(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_BaseItm
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_BaseItmNm
	frm1.vspdData.Text = arrRet(1)		

End Function
'------------------------------------------  SetWtUnit()  --------------------------------------------------
'	Name : SetWtUnit()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetWtUnit(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_UnitOfWeight
	frm1.vspdData.Text = arrRet(0)
End Function

'------------------------------------------  SetWtUnit()  --------------------------------------------------
'	Name : SetWtUnit()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetGrossUnit(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_UnitOfGrossWeight
	frm1.vspdData.Text = arrRet(0)
End Function
'------------------------------------------  SetHsCd()  --------------------------------------------------
'	Name : SetHsCd()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetHsCd(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_HsCd
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_HsUnit
	frm1.vspdData.Text = arrRet(3)		
End Function

'------------------------------------------  SetBillHdr()  -----------------------------------------------
'	Name : SetBillHdr()
'	Description : Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetBillHdr(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_VatType
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_VatRate
	frm1.vspdData.Text = arrRet(2)	
	
End Function

'==========================================================================================
'   Event Name : txtItemCd1_onChange()
'   Event Desc :
'==========================================================================================

Sub txtItemCd1_onChange()
	With frm1
		If .txtItemCd1.value = "" Then
			.txtItemNm1.value = ""
			.txtItemSpec1.value = ""
			.txtItemProcType1.value = ""
			
			.txtItemCd1.focus
			Set gActiveElement = document.activeElement
		Else	
			Call LookUpItemByPlant()
		End If
	End With
End Sub

'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'	Name : LookUpItem ByPlant()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 
Function LookUpItemByPlant()
	Dim iStrWhereSQL
	Dim strITEM_CD
	Dim strITEM_NM
	Dim strSPEC
	Dim strPROCUR_TYPE_CD
	Dim strPROCUR_TYPE_NM

    iStrWhereSQL = "A.ITEM_CD = B.ITEM_CD AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd1.value, "''", "S") & " AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	Call CommonQueryRs(" A.ITEM_CD, A.ITEM_NM, A.SPEC, B.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", B.PROCUR_TYPE) "," B_ITEM A, B_ITEM_BY_PLANT B ",iStrWhereSQL ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	strITEM_CD = lgF0
	strITEM_NM = lgF1
	strSPEC = lgF2
	strPROCUR_TYPE_CD = lgF3
	strPROCUR_TYPE_NM = lgF4
		
	strITEM_CD			=	replace(strITEM_CD,Chr(11),"")
	strITEM_NM			=	replace(strITEM_NM,Chr(11),"")
	strSPEC				=	replace(strSPEC,Chr(11),"")
	strPROCUR_TYPE_CD	=	replace(strPROCUR_TYPE_CD,Chr(11),"")
	strPROCUR_TYPE_NM	=	replace(strPROCUR_TYPE_NM,Chr(11),"")
	
	frm1.txtItemNm1.value = strITEM_NM
	frm1.txtItemSpec1.value = strSPEC
	frm1.txtItemProcType1.value = strPROCUR_TYPE_NM
	frm1.htxtItemProcType1.value = strPROCUR_TYPE_CD
End Function

'------------------------------------------  ChkBtnAll()  --------------------------------------------------
'	Name : ChkBtnAll()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function btnSelectAll_Clicked()
	Dim LngRow
	
	If frm1.vspdData.MaxRows <= 0 Then Exit Function
	
	lgFlgBtnSelectAllClicked = True
	frm1.btnSelectAll.disabled = True
	With frm1.vspdData
		
		.ReDraw = False
		If lgFlgAllSelected = False Then 'select all clicked
				
			For LngRow = 1 To .MaxRows
				Call .SetText(C_Select,LngRow,"1")
				Call SetSpreadUnLock(C_Select, LngRow)
				If lgInsrtFlg <> True Then
					ggoSpread.UpdateRow LngRow
				End If
			Next

			Call InitData(1,1)	
			
			frm1.btnSelectAll.value = "��ü�������"
			lgFlgAllSelected = True
			
		Else 'deselect all clicked
			For LngRow = .MaxRows To 1 Step -1
				If GetSpreadText(frm1.vspdData,C_Select,LngRow,"X","X") = "1" _ 
				And GetSpreadText(frm1.vspdData,0,LngRow,"X","X") <> ggoSpread.InsertFlag Then
					Call .SetText(C_Select,LngRow,"0")
					Call ggoSpread.EditUndo(LngRow, LngRow)
					Call SetSpreadLock1(C_Select, LngRow)
				End If
			Next
			
			Call InitData(1,1)

			frm1.btnSelectAll.value = "��ü����"
			lgFlgAllSelected = False
		End If
		.ReDraw = True
	End With
	
	frm1.btnSelectAll.disabled = False
	lgFlgBtnSelectAllClicked = False
	
End Function

'==========================================================================================
'   Event Name : LookUpVatType
'   Event Desc : �ΰ���Ÿ�� ������ ����Ǿ����� �ΰ����� ��� 
'==========================================================================================
Sub LookUpVatType(ByVal VatType)
	Dim strVal

    Err.Clear                                                               <%'��: Protect system from crashing%>

	If VatType = "" Then Exit Sub
	If   LayerShowHide(1) = False Then Exit Sub

	strVal = BIZ_PGM_LOOKUPVATTYPE_ID & "?txtMode=" & parent.UID_M0001								<%'��: �����Ͻ� ó�� ASP�� ���� %>
	strVal = strVal & "&txtVatType=" & Trim(VatType)					<%'��: ��ȸ ���� ����Ÿ %>
	strVal = strVal & "&Row=" & frm1.vspdData.Row										'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)												<%'��: �����Ͻ� ASP �� ���� %>
End Sub

Function LookUpVatTypeok(ByVal VatType,ByVal VatRate, ByVal Row)
	With frm1.vspdData
		.ReDraw = False
		.Row = CLng(Row)
		
		.Col = C_VatType
		.Text = Trim(VatType)
		.Col = C_VatRate
		.Text = VatRate
			
		.ReDraw = True
	End With
	IsOpenPop = False
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
	Err.Clear
	
	iDBSYSDate = "<%=GetSvrDate%>"											'��: DB�� ���� ��¥�� �޾ƿͼ� ���۳�¥�� ����Ѵ�.
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	
	Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","2")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
	
	'----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet															'��: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal	'�Լ� ���ǰ� ���� 
	Call InitVariables	'�Լ� ���ǰ� ����											'��: Initializes local global variables
	
	Call SetToolbar("11000101001011")												'��: ��ư ���� ���� 
	
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtPlantCd1.value = parent.gPlant
			frm1.txtPlantNm1.value = parent.gPlantNm
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
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	Call SetPopupMenuItemInf("1001111111") 
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		ggoSpread.Source = frm1.vspdData
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================
Sub vspdData_DblClick(ByVal Col , ByVal Row )
    If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.Col = C_Select
		.Row = Row

		If .Value = "1" Then
			ggoSpread.UpdateRow Row
		End if

		If Col = C_VatType Then
			.Row = Row
			.Col = C_VatType

			Call LookUpVatType(.Text)
		End If			
		
		If Col = C_PrcCtrlIndNm Then
		   Call vspdData_ComboSelChange (C_PrcCtrlIndNm,Row)
		End If   
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	'----------  Coding part  -------------------------------------------------------------   

	If frm1.vspdData.Row <= 0 Or lgFlgBtnSelectAllClicked = True Then Exit Sub
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		If gMouseClickStatus = "SPC" Or lgFlgCancelClicked = True Then
			If Col = C_Select And Not (lgFlgCopyClicked) Then
				If GetSpreadText(frm1.vspdData,C_Select,Row,"X","X") = "0" Then
					.Redraw = false
					Call SetSpreadLock1(C_Select, Row)
					Call ggoSpread.EditUndo(Row, Row)
					Call InitData(1,1)
					.Redraw = true
				Else
					.Redraw = false
					Call SetSpreadUnLock(C_Select, Row)	
					.Redraw = true
				End If
			End If
		End If
				
		Select Case Col
			Case C_UnitPopup
				Call OpenUnitPopup(GetSpreadText(frm1.vspdData,C_Unit,Row,"X","X"))
				
			Case C_ItmGroupPopup
				Call OpenItemGroupPopup(GetSpreadText(frm1.vspdData,C_ItmGroupCd,Row,"X","X"))
				
			Case C_BaseItmPopup
				Call OpenBaseItemPopup(GetSpreadText(frm1.vspdData,C_BaseItmPopup,Row,"X","X"))
				
			Case C_WeightUnitPopup
				
				Call OpenWtUnitPopup(GetSpreadText(frm1.vspdData,C_WeightUnitPopup,Row,"X","X"))
			
			Case C_GrossUnitPopup
				Call OpenGrossUnit(GetSpreadText(frm1.vspdData,C_GrossUnitPopup,Row,"X","X"))
				
			Case C_HsCdPopup
				Call OpenHsPopup(GetSpreadText(frm1.vspdData,C_HsCdPopup,Row,"X","X"))
				
			Case C_VatTypePopup
				Call OpenBillHdr(GetSpreadText(frm1.vspdData,C_VatTypePopup,Row,"X","X"))
				
			Case C_Select
				If lgInsrtFlg <> True Then
					If Buttondown = 1 Then
						ggoSpread.Source = frm1.vspdData
						ggoSpread.UpdateRow Row
					Else
						ggoSpread.Source = frm1.vspdData
					End If
				End If
				
		End Select
    End With
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
			Case  C_ItmAcc
				.Col = Col
				intIndex = .Value
				.Col = C_HdnItmAcc
				.Value = intIndex
			Case  C_SumItmClass
				.Col = Col
				intIndex = .Value
				.Col = C_HdnSumItmClass
				.Value = intIndex
			Case  C_PrcCtrlIndNm
				.Col = Col
				intIndex = .value
				.Col = C_PrcCtrlInd
				.value = intIndex
				If Trim(.Text) = "M" Then
					.Col = C_UnitPrice
					.Text = "0"
					ggoSpread.SpreadLock		C_UnitPrice,		Row, C_UnitPrice,		Row
					ggoSpread.SSSetProtected 	C_UnitPrice, 		Row, Row
				Else
					ggoSpread.SpreadUnLock		C_UnitPrice,		Row, C_UnitPrice,		Row
					ggoSpread.SSSetRequired 	C_UnitPrice, 		Row, Row
				End If
			Case  C_PrcCtrlInd
				.Col = Col
				intIndex = .value
				.Col = C_PrcCtrlIndNm
				.value = intIndex						
		End Select
    
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row >= NewRow Then Exit Sub
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
    If OldLeft <> NewLeft Then Exit Sub
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)   ': Query ��ư�� disable ��Ŵ.
            
            frm1.vspdData.ReDraw = False
            If DBQuery = False Then 
               Call RestoreToolBar()
               frm1.vspdData.ReDraw = True
               Exit Sub
            End If 
            frm1.vspdData.ReDraw = True
		End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : txtPlantCd_OnChange
'   Event Desc : This function is Setting the txtPlantCd,txtPlantNm
'==========================================================================================
Sub txtPlantCd_OnBlur()
	With frm1
		If Trim(.txtPlantCd.value) = "" Then
			.txtPlantNm.value = ""
			.txtPlantCd1.value = ""
			.txtPlantNm1.value = ""
		Else
			.txtPlantCd1.value = UCase(Trim(.txtPlantCd.value))
			.txtPlantNm1.value = UCase(Trim(.txtPlantNm.value))
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
    Dim strPlantCd
    Dim strPlantNm
    Dim strPlantItem
    Dim strPlantItemNm
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			'��: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If

	If frm1.txtPlantCd1.value = "" Then
		frm1.txtPlantNm1.value = ""
	Else
		strPlantCd = frm1.txtPlantCd1.value 
		strPlantNm = frm1.txtPlantNm1.value 
	End If
	
	If frm1.txtItemCd1.value = "" Then
		frm1.txtItemNm1.value = ""
	Else
		strPlantItem = frm1.txtItemCd1.value 
		strPlantItemNm = frm1.txtItemNm1.value 
	End If

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field

	If strPlantCd <> "" Then    
		frm1.txtPlantCd1.value = strPlantCd
		frm1.txtPlantNm1.value = strPlantNm
	End If
	
	If strPlantItem <> "" Then    
		frm1.txtItemCd1.value = strPlantItem
		frm1.txtItemNm1.value = strPlantItemNm
	End If

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
    End If     																'��: Query db data

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
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '��: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If frm1.txtItemCd1.value = "" Then
		frm1.txtItemNm1.value = ""
	End If
	    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Precheck area
    '-----------------------
   	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     							                                      '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear
    
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	FncCopy = False
    lgInsrtFlg = True
	lgFlgCopyClicked = True
	ggoSpread.Source = frm1.vspdData
	
	
	With frm1.vspdData
		.ReDraw = False
		If .ActiveRow > 0 Then
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
			
			.EditMode = True
			.Focus
		End If
		.ReDraw = True

		'-------------------------------------------------
		' Default Value Setting
		'-------------------------------------------------
		ggoSpread.SpreadUnLock	C_Select, .ActiveRow, C_Select, .ActiveRow
		.Col = C_Select
		.Row = .ActiveRow
		.value = 1
		
		.Col = C_Item
		.Row = .ActiveRow
		.value = ""
		
		.Col = C_PrcCtrlInd
		.Row = .ActiveRow
		
		If Trim(.text) = "M" Then
			.Col = C_UnitPrice
			.Row = .ActiveRow
			.Text = "0"

			ggoSpread.SpreadLock	C_UnitPrice, .ActiveRow, C_UnitPrice, .ActiveRow
			ggoSpread.SSSetProtected C_UnitPrice, .ActiveRow, .ActiveRow
			
		Else
			ggoSpread.SpreadUnLock	C_UnitPrice, .ActiveRow, C_UnitPrice, .ActiveRow
			ggoSpread.SSSetRequired	C_UnitPrice, .ActiveRow, .ActiveRow
		End If
		
		
	End With
	
	lgInsrtFlg = False	
	lgFlgCopyClicked = False
	
	Set gActiveElement = document.activeElement
	
	If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If
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
	If frm1.vspdData.MaxRows <= 0 Then Exit Function
	
	lgFlgCancelClicked = True
	
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.ReDraw = False
		Call ggoSpread.EditUndo(.ActiveRow,.ActiveRow)
		Call InitData(1,1)
	'	Call SetSpreadLock1(C_Select, .ActiveRow)
		.ReDraw = True
	End With
	
	lgFlgCancelClicked = False
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
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()	
	Call InitData(1,1)
End Sub


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow
	Dim newRow
    
    On Error Resume Next	                                                     '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False 
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If
    
	lgInsrtFlg = True
	
    With frm1.vspdData
		
		.focus		
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow , imRow
		SetSpreadColor .ActiveRow, .ActiveRow + imRow - 1
		
    	'--------------------------------------------
    	' Default Setting 
    	' �߰��� Row������ŭ �÷��� �ʱ�ȭ ��Ų��.
    	'--------------------------------------------    	
	    For newRow = 0 To Cint(imRow) - 1
	    	ggoSpread.SpreadUnLock	C_Select, .ActiveRow + newRow, C_Select, .ActiveRow + newRow
	    	.Col = C_Select
	    	.Row = .ActiveRow + newRow
			.value = 1
				
			.Col = C_Phantom
			.Row = .ActiveRow + newRow
			.Text = "N"
				
			.Col = C_BlanketPur
			.Row = .ActiveRow + newRow
			.Text = "N"

			.Col = C_DefaultFlg
			.Row = .ActiveRow + newRow
			.Text = "Y"
				
			.Col = C_PicFlg
			.Row = .ActiveRow + newRow
			.Text = "N"
				
			.Col = C_IBPValidFromDt
			.Row = .ActiveRow + newRow
			.Text = StartDate
				
			.Col = C_IBPValidToDt
			.Row = .ActiveRow + newRow
			.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")

			.Col = C_ValidFromDt
			.Row = .ActiveRow + newRow
			.Text = StartDate
				
			.Col = C_ValidToDt
			.Row = .ActiveRow + newRow
			.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
				
			.Col = C_PrcCtrlInd
			.Row = .ActiveRow + newRow
			.Text = "S"

			Call vspdData_ComboSelChange(C_PrcCtrlInd, .ActiveRow + newRow)
		Next
		.ReDraw = True
	    .EditMode = True
		
	End With
	
	Set gActiveElement = document.activeElement
	lgInsrtFlg = False
	
	If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
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
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         '��: Processing is NG
    
    Call LayerShowHide(1)
    
    Dim strVal

    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001								'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)				'��: ��ȸ ���� �� 
		strVal = strVal & "&cboItemAccount=" & Trim(frm1.hItemAccount.value)				'��: ��ȸ ���� �� 
		strVal = strVal & "&cboItemClass=" & Trim(frm1.hItemClass.value)				'��: ��ȸ ���� �� 
		strVal = strVal & "&rdoPhantomFlg=" & Trim(frm1.hPhantomFlg.value)				'��: ��ȸ ���� �� 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtHighItemGroupCd.value)				'��: ��ȸ ���� �� 
		strVal = strVal & "&cboItemAccount=" & Trim(frm1.cboItemAcct.value)				'��: ��ȸ ���� �� 
		strVal = strVal & "&cboItemClass=" & Trim(frm1.cboItemClass.value)				'��: ��ȸ ���� �� 
		If frm1.rdoPhantomFlg1.checked = True Then 
			strVal = strVal & "&rdoPhantomFlg=" & Trim(frm1.rdoPhantomFlg1.value)				'��: ��ȸ ���� �� 
		ElseIf frm1.rdoPhantomFlg2.checked = True Then
			strVal = strVal & "&rdoPhantomFlg=" & Trim(frm1.rdoPhantomFlg2.value)				'��: ��ȸ ���� �� 
		Else
			strVal = strVal & "&rdoPhantomFlg=" & Trim(frm1.rdoPhantomFlg3.value)				'��: ��ȸ ���� �� 
		End If
		strVal = strVal & "&txtItemCd1=" & Trim(frm1.txtItemCd1.value)				'��: ��ȸ ���� ����Ÿ	
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    End If

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(LngMaxRow)													'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
    Call InitData(LngMaxRow,1)
	
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field
    
    Call SetToolbar("11000101001111")
    
    frm1.btnCopy.disabled = True
	frm1.btnSelectAll.disabled = True
	frm1.btnSelectAll.value = "��ü����"
	lgFlgAllSelected = False
		
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim IntRows 
    Dim lGrpcnt 
    Dim strVal
    Dim GenVal
	Dim IntRetCD
	Dim iColSep
	Dim TmpBuffer
	Dim iTotalStr

	DbSave = False														'��: Processing is NG
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value  = parent.gUsrID
		
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = Parent.gColSep
    lGrpCnt = 1
    ReDim TmpBuffer(0)
	
	If frm1.rdoProdWorkSet1.checked = True Then
		GenVal = "10000000"
	ElseIf frm1.rdoProdWorkSet2.checked = True Then
		GenVal = "11000000"
	ElseIf frm1.rdoProdWorkSet3.checked = True Then
		GenVal = "11100000"
	ElseIf frm1.rdoProdWorkSet4.checked = True Then
		GenVal = "11110000"
	ElseIf frm1.rdoProdWorkSet5.checked = True Then
		GenVal = "11111000"
	End If
	
	With frm1.vspdData

		For IntRows = 1 To .MaxRows

			.Row = IntRows
			.Col = 0
			
			strVal = ""
			
			If .Text = ggoSpread.InsertFlag Then
				strVal = strVal & "I" & iColSep & IntRows & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
			ElseIf .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "C" & iColSep	& IntRows & iColSep				'��: U=Update
			End If

			Select Case .Text
		    
			    Case ggoSpread.InsertFlag
					
			        .Col = C_Item								'2
			        strVal = strVal & Trim(.Text) & iColSep
			            
			        .Col = C_ItmNm								'3
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_ItmFormalNm						'4
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_HdnItmAcc							'5
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_Unit								'6
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_ItmGroupCd							'7
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_Phantom							'8
			        strVal = strVal & Trim(.Text) & iColSep
    
			        .Col = C_BlanketPur							'9
			        strVal = strVal & Trim(.Text) & iColSep
					
					.Col = C_BaseItm							'10
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_HdnSumItmClass						'11
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_DefaultFlg							'12
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_ItmSpec							'13
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_UnitWeight							'14
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_UnitOfWeight						'15
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_DrawNo								'16
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_HsCd								'17
			        strVal = strVal & Trim(.Text) & iColSep
					
					.Col = C_HsUnit								'18
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_IBPValidFromDt						'19
					strVal = strVal & Trim(.Text) & iColSep						
				    
			        .Col = C_IBPValidToDt						'20
					strVal = strVal & Trim(.Text) & iColSep

					.Col = C_PrcCtrlInd							'21
					strVal = strVal & Trim(.Text) & iColSep		

					If (Trim(UCase(frm1.htxtItemProcType1.value)) = "M" And Trim(UCase(.Text)) = "M") Or (Trim(UCase(frm1.htxtItemProcType1.value)) = "O" And Trim(UCase(.Text)) = "M") Then
						IntRetCD = DisplayMsgBox("122726", "X", "X", "X")	'���ޱ����� �系����ǰ�̸� �ܰ������� ǥ�شܰ��� �����մϴ�.
						Call LayerShowHide(0)
						Exit Function
					End If
										
					.Col = C_UnitPrice							'22
					strVal = strVal & Trim(.Text) & iColSep		

					.Col = C_VatType							'23
					strVal = strVal & Trim(.Text) & iColSep		
					
					.Col = C_VatRate							'24
					strVal = strVal & Trim(.Text) & iColSep		
														
					strVal = strVal & GenVal & iColSep			'25
					
					.Col = C_UnitGrossWeight					'26
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_UnitOfGrossWeight					'27
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_CBM								'28
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_CBMDesc							'29									'��: ������ ����Ÿ�� Row �и���ȣ�� �ִ´�		        
			        strVal = strVal & Trim(.Text) & parent.gRowSep
					
					ReDim Preserve TmpBuffer(lGrpCnt - 1)
					TmpBuffer(lGrpCnt - 1) = strVal
					
			        lGrpCnt = lGrpCnt + 1
			        
			    Case ggoSpread.UpdateFlag

					.Col = C_Item								'2
			        strVal = strVal & Trim(.Text) & iColSep

					.Col = C_PrcCtrlInd							'3
					strVal = strVal & Trim(.Text) & iColSep		

					If (Trim(UCase(frm1.htxtItemProcType1.value)) = "M" And Trim(UCase(.Text)) = "M") Or (Trim(UCase(frm1.htxtItemProcType1.value)) = "O" And Trim(UCase(.Text)) = "M") Then
						IntRetCD = DisplayMsgBox("122726", "X", "X", "X")	'���ޱ����� �系����ǰ�̸� �ܰ������� ǥ�شܰ��� �����մϴ�.
						Call LayerShowHide(0)
						Exit Function
					End If
					
					.Col = C_UnitPrice							'4
					strVal = strVal & Trim(.Text) & iColSep		

					.Col = C_Phantom							'5
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_IBPValidFromDt						'6
					strVal = strVal & Trim(.Text) & iColSep						
				    
			        .Col = C_IBPValidToDt						'7
					strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_ValidFromDt						'8
					strVal = strVal & Trim(.Text) & iColSep						
				    
			        .Col = C_ValidToDt						'9
					strVal = strVal & Trim(.Text) & iColSep

					strVal = strVal & GenVal & parent.gRowSep			'10			'��: ������ ����Ÿ�� Row �и���ȣ�� �ִ´�		        

					ReDim Preserve TmpBuffer(lGrpCnt - 1)
					TmpBuffer(lGrpCnt - 1) = strVal
					
					lGrpcnt = lGrpcnt + 1             
			End Select
	   Next
	End With
	
	iTotalStr = Join(TmpBuffer)
	frm1.txtMaxRows.value = lGrpCnt-1										'��: Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread.value = iTotalStr											'��: Spread Sheet ������ ���� 

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: ���� �����Ͻ� ASP �� ���� 

    DbSave = True                                                           '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables

    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.MaxRows = 0

    Call FncQuery()
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>
<!--'#########################################################################################################
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ǰ������COPY</font></td>
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
									<TD CLASS=TD6 NOWRAP>
										<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=7 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="24"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value,0">&nbsp;
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHighItemGroupCd" SIZE=8 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btHighItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()" >&nbsp;<INPUT TYPE=TEXT NAME="txtHighItemGroupNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="ǰ�����" STYLE="Width: 160px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����� ǰ��Ŭ����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemClass" ALT="����� ǰ��Ŭ����" STYLE="Width: 160px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>Phantom ����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPhantomFlg" ID="rdoPhantomFlg1" CLASS="RADIO" tag="11" Value="A" CHECKED><LABEL FOR="rdoPhantomFlg1">��ü</LABEL>
														 <INPUT TYPE="RADIO" NAME="rdoPhantomFlg" ID="rdoPhantomFlg2" CLASS="RADIO" tag="11" Value="Y"><LABEL FOR="rdoPhantomFlg2">��</LABEL>
									     				 <INPUT TYPE="RADIO" NAME="rdoPhantomFlg" ID="rdoPhantomFlg3" CLASS="RADIO" tag="11" Value="N"><LABEL FOR="rdoPhantomFlg3">�ƴϿ�</LABEL></TD>
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
								<TD CLASS=TD5 NOWRAP>���ذ���</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd1" SIZE=7 MAXLENGTH=4 tag="24" ALT="���ذ���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant1()">&nbsp;
									<INPUT TYPE=TEXT NAME="txtPlantNm1" SIZE=20 tag="24">
								</TD>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoProdWorkSet" ID="rdoProdWorkSet1" tag="21" CHECKED><LABEL FOR="rdoProdWorkSet1">���庰 ǰ�� �������� ����</LABEL>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ǰ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="22XXXU" ALT="����ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd1.value,1"></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoProdWorkSet" ID="rdoProdWorkSet2" tag="21"><LABEL FOR="rdoProdWorkSet2">BOM �������� ����</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ǰ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=40 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoProdWorkSet" ID="rdoProdWorkSet3" tag="21"><LABEL FOR="rdoProdWorkSet3">Routing �������� ����</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ǰ��԰�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec1" SIZE=40 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoProdWorkSet" ID="rdoProdWorkSet4" tag="21"><LABEL FOR="rdoProdWorkSet4">��ǰ�� ���� �������� ����</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ǰ�����ޱ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemProcType1" SIZE=40 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoProdWorkSet" ID="rdoProdWorkSet5" tag="21"><LABEL FOR="rdoProdWorkSet5">Bill Of Resource �������� ����</LABEL></TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="22" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnCopy" CLASS="CLSMBTN" Flag=1 ONCLICK="FncSave">COPY</BUTTON>&nbsp;
						<BUTTON NAME="btnSelectAll" CLASS="CLSMBTN" Flag=1 ONCLICK="btnSelectAll_Clicked">��ü����</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TabIndex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemAccount" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hPhantomFlg" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemClass" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="htxtItemProcType1" tag="24" TabIndex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</BODY>
</HTML>
