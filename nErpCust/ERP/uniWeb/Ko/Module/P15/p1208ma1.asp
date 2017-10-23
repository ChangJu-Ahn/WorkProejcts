
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Rouing basic infomation
'*  3. Program ID           : p1208ma1
'*  4. Program Name         : Manufacturing Instruction
'*  5. Program Desc         : Entry Manufacturing Instruction
'*  6. Component List       : Using HR ADO CUD Source.
'*  7. Modified date(First) : 2002/03/20
'*  8. Modified date(Last)  : 2002/12/17
'*  9. Modifier (First)     : Chen, Jae Hyun
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_LOOKUP_ID	= "p1208mb0.asp"								' Lookup Item By Plant

'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p1208mb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p1208mb2.asp"								'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID	= "p1208mb3.asp"
Const BIZ_PGM_COPY_ID	= "p1208mb4.asp"								'��: ǥ���۾����� ���� �����Ͻ� ���� ASP�� 

' Grid 1(vspdData1) - Operation 
Dim C_OprNo
Dim C_WcCd
Dim C_WcNm
Dim C_JobCd
Dim C_JobCdNm
Dim C_InsideFlg
Dim C_MfgLt
Dim C_QueueLT
Dim C_SetupLT
Dim C_WaitLT
Dim C_FixRunTime
Dim C_OprLT
Dim C_RuntimeQty
Dim C_RuntimeUnit
Dim C_MoveLT
Dim C_OverlapOpr
Dim C_OverlapLT
Dim C_BpCd
Dim C_BpNm
Dim C_CurCd
Dim C_SubcontractPrc
Dim C_Milestone
Dim C_InspFlg
Dim C_RoutOrder
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_TaxType

' Grid 2(vspdData2) - Operation
Dim C_Seq2
Dim C_WICd2
Dim C_WICdPopup2
Dim C_WIDesc2
Dim C_ValidStartDt2
Dim C_ValidEndDt2
Dim C_PlantCd2
Dim C_ItemCd2
Dim C_RoutingNo2
Dim C_OprNo2
Dim C_HdnSeq2

' Grid 3(vspdData3) - Hidden
Dim C_Seq3
Dim C_WICd3
Dim C_WICdPopup3
Dim C_WIDesc3
Dim C_ValidStartDt3
Dim C_ValidEndDt3
Dim C_PlantCd3
Dim C_ItemCd3
Dim C_RoutingNo3
Dim C_OprNo3
Dim C_HdnSeq3

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgCurrRow
Dim lgFlgQueryCnt

Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgBlnFlgSaveValue
Dim lgBlnFlgLookupValue
Dim lgBlnFlgQryValue
Dim lgLngCnt
Dim lgOldRow
Dim lgRow         

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvGridId)

	If pvGridId = "*" Or pvGridId = "A" Then
		' Grid 1(vspdData1) - Operation 
		C_OprNo				= 1
		C_WcCd				= 2
		C_WcNm				= 3
		C_JobCd				= 4
		C_JobCdNm			= 5
		C_InsideFlg			= 6
		C_MfgLt				= 7
		C_QueueLT			= 8
		C_SetupLT			= 9
		C_WaitLT			= 10
		C_FixRunTime        = 11
		C_OprLT				= 12
		C_RuntimeQty		= 13
		C_RuntimeUnit		= 14
		C_MoveLT			= 15
		C_OverlapOpr		= 16
		C_OverlapLT			= 17
		C_BpCd				= 18
		C_BpNm				= 19
		C_CurCd				= 10
		C_SubcontractPrc	= 21
		C_Milestone			= 22
		C_InspFlg			= 23
		C_RoutOrder			= 24
		C_ValidFromDt		= 25
		C_ValidToDt			= 26
		C_TaxType			= 27
	End If
	
	If pvGridId = "*" Or pvGridId = "B" Then
		' Grid 2(vspdData2) - Operation
		C_Seq2				= 1
		C_WICd2				= 2
		C_WICdPopup2		= 3
		C_WIDesc2			= 4
		C_ValidStartDt2		= 5
		C_ValidEndDt2		= 6
		C_PlantCd2			= 7
		C_ItemCd2			= 8
		C_RoutingNo2		= 9
		C_OprNo2			= 10
		C_HdnSeq2			= 11
	End If
	
	If pvGridId = "*"  Or pvGridId = "C" Then
		' Grid 3(vspdData3) - Hidden
		C_Seq3				= 1
		C_WICd3				= 2
		C_WICdPopup3		= 3
		C_WIDesc3			= 4
		C_ValidStartDt3		= 5
		C_ValidEndDt3		= 6
		C_PlantCd3			= 7
		C_ItemCd3			= 8
		C_RoutingNo3		= 9
		C_OprNo3			= 10
		C_HdnSeq3			= 11
	End If
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgRow = 0
    lgBlnFlgSaveValue = False
	lgBlnFlgLookupValue = False
	lgBlnFlgQryValue = False	
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtStdDt.Text = StartDate
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvGridId)

	Call InitSpreadPosVariables(pvGridId)

	If pvGridId = "*" Or pvGridId = "A" Then

		With frm1.vspdData1

			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

			.ReDraw = False

			.MaxCols = C_TaxType + 1
			.MaxRows = 0

			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetEdit	C_OprNo, "����", 8
			ggoSpread.SSSetEdit	C_WcCd, "�۾���", 12				
			ggoSpread.SSSetEdit	C_WcNm, "�۾����", 12
			ggoSpread.SSSetEdit	C_JobCd, "�۾�", 8
			ggoSpread.SSSetEdit	C_JobCdNm, "�۾���", 15
			ggoSpread.SSSetEdit	C_InsideFlg, "����Ÿ��", 8
			ggoSpread.SSSetEdit	C_MfgLT,	 "����L/T", 7, 1
			ggoSpread.SSSetTime C_QueueLT, "Queue�ð�",	10, 2 ,1 ,1
			ggoSpread.SSSetTime C_SetupLT, "��ġ�ð�",	10, 2 ,1 ,1
			ggoSpread.SSSetTime C_WaitLT, "���ð�",	10, 2 ,1 ,1
			ggoSpread.SSSetTime C_FixRunTime, "���������ð�", 10, 2 ,1 ,1
			ggoSpread.SSSetTime C_OprLT, "���������ð�", 10, 2 ,1 ,1
			ggoSpread.SSSetFloat	C_RuntimeQty, "���ؼ���", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit C_RuntimeUnit, "���ش���", 6
			ggoSpread.SSSetTime C_MoveLT, "�̵��ð�",	10,2 ,1 ,1		
			ggoSpread.SSSetEdit	C_OverlapOpr,			"Overlap ����", 7
			ggoSpread.SSSetEdit	C_OverlapLt,			"Overlap L/T", 8, 1
			ggoSpread.SSSetEdit	C_BpCd,					"����ó", 10
			ggoSpread.SSSetEdit	C_BpNm,					"����ó��", 15
			ggoSpread.SSSetEdit	C_CurCd,				"��ȭ", 6,,,3,2
			'ggoSpread.SSSetFloat	C_SubcontractPrc,	"�������ִܰ�",15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SubcontractPrc,	"�������ִܰ�",15,"C"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit	C_Milestone,			"Milestone", 10
			ggoSpread.SSSetEdit	C_InspFlg,				"�˻翩��", 10
			ggoSpread.SSSetEdit	C_RoutOrder,			"�����ܰ�", 10
			ggoSpread.SSSetDate C_ValidFromDt,			"������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate C_ValidToDt,			"������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit	C_TaxType,				"VAT����", 15,,,20,2	
	
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SSSetSplit2(2)
	
			.ReDraw = True
    
		End With
	End If
	
	If pvGridId = "*" Or pvGridId = "B" Then

		With frm1.vspdData2
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

			.ReDraw = False

			.MaxCols = C_HdnSeq2 + 1
			.MaxRows = 0

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit		C_Seq2,				"�۾�����", 8,,,3,2	
			ggoSpread.SSSetEdit		C_WICd2,			"�����۾�", 14,,,10,2
			ggoSpread.SSSetButton 	C_WICdPopup2
			ggoSpread.SSSetEdit		C_WIDesc2,			"�����۾�����", 72
			ggoSpread.SSSetDate 	C_ValidStartDt2,	"��ȿ������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ValidEndDt2,		"��ȿ������", 11, 2, parent.gDateFormat	
			ggoSpread.SSSetEdit 	C_PlantCd2, 		"����", 7
			ggoSpread.SSSetEdit 	C_ItemCd2,			"ǰ��", 7
			ggoSpread.SSSetEdit 	C_RoutingNo2, 		"�����", 7
			ggoSpread.SSSetEdit 	C_OprNo2, 			"����", 7
			ggoSpread.SSSetEdit 	C_HdnSeq2, 			"�۾�����", 7
				
			Call ggoSpread.MakePairsColumn(C_WICd2, C_WICdPopup2)
			Call ggoSpread.SSSetColHidden(C_PlantCd2, C_HdnSeq2, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SSSetSplit2(3)	
	
			.ReDraw = True
	
		End With
	End If
	
	If pvGridId = "*" Or pvGridId = "C" Then

		With frm1.vspdData3
			
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

			.ReDraw = False

			.MaxCols = C_HdnSeq3 + 1
			.MaxRows = 0

			Call GetSpreadColumnPos("C")
			ggoSpread.SSSetEdit		C_Seq3, "�۾�����", 8,,,3,2	
			ggoSpread.SSSetEdit		C_WICd3, "�����۾�", 14,,,10,2
			ggoSpread.SSSetButton 	C_WICdPopup3
			ggoSpread.SSSetEdit		C_WIDesc3, "�����۾�����", 50
			ggoSpread.SSSetDate 	C_ValidStartDt3,			"��ȿ������", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ValidEndDt3,			"��ȿ������", 11, 2, parent.gDateFormat	
			ggoSpread.SSSetEdit 	C_PlantCd3, 	"����", 7
			ggoSpread.SSSetEdit 	C_ItemCd3, 	"ǰ��", 7
			ggoSpread.SSSetEdit 	C_RoutingNo3, 	"�����", 7
			ggoSpread.SSSetEdit 	C_OprNo3, 	"����", 7
			ggoSpread.SSSetEdit 	C_HdnSeq3, 	"�۾�����", 7		
	
			.ReDraw = True
    
		End With
	End If

	Call SetSpreadLock(pvGridId)
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadLock(ByVal pvGridId)

    With frm1

	If pvGridId = "*" Or pvGridId = "A" Then
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
	
	If pvGridId = "*" Or pvGridId = "B" Then
		ggoSpread.Source = frm1.vspdData2
		.vspdData2.ReDraw = False
		ggoSpread.SpreadLock C_Seq2, -1, C_Seq2
		ggoSpread.SpreadLock C_WIDesc2, -1, C_WIDesc2
		ggoSpread.SpreadLock C_ValidStartDt2, -1, C_ValidStartDt2
		ggoSpread.SpreadLock C_ValidEndDt2, -1, C_ValidEndDt2
		ggoSpread.SSSetRequired	 C_WICd2, -1
		ggoSpread.SSSetProtected .vspdData2.MaxCols, -1
		.vspdData2.ReDraw = True
	End If
   
    End With

End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData2
		.Redraw = False
		ggoSpread.Source = frm1.vspdData2    
		ggoSpread.SpreadUnLock   C_Seq2,	pvStartRow, C_Seq2, pvEndRow
		ggoSpread.SSSetRequired  C_Seq2,	pvStartRow, pvEndRow    
		ggoSpread.SSSetRequired	C_WICd2,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_WIDesc2,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ValidStartDt2,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ValidEndDt2,		pvStartRow, pvEndRow
		.Redraw = True
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		' Grid 1(vspdData1) - Operation 
		Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_OprNo				= iCurColumnPos(1)
			C_WcCd				= iCurColumnPos(2)
			C_WcNm				= iCurColumnPos(3)
			C_JobCd				= iCurColumnPos(4)
			C_JobCdNm			= iCurColumnPos(5)
			C_InsideFlg			= iCurColumnPos(6)
			C_MfgLt				= iCurColumnPos(7)
			C_QueueLT			= iCurColumnPos(8)
			C_SetupLT			= iCurColumnPos(9)
			C_WaitLT			= iCurColumnPos(10)
			C_FixRunTime        = iCurColumnPos(11) 
			C_OprLT				= iCurColumnPos(12)
			C_RuntimeQty		= iCurColumnPos(13)
			C_RuntimeUnit		= iCurColumnPos(14)
			C_MoveLT			= iCurColumnPos(15)
			C_OverlapOpr		= iCurColumnPos(16)
			C_OverlapLT			= iCurColumnPos(17)
			C_BpCd				= iCurColumnPos(18)
			C_BpNm				= iCurColumnPos(19)
			C_CurCd				= iCurColumnPos(20)
			C_SubcontractPrc	= iCurColumnPos(21)
			C_Milestone			= iCurColumnPos(22)
			C_InspFlg			= iCurColumnPos(23)
			C_RoutOrder			= iCurColumnPos(24)
			C_ValidFromDt		= iCurColumnPos(25)
			C_ValidToDt			= iCurColumnPos(26)
			C_TaxType			= iCurColumnPos(27)

		' Grid 2(vspdData2) - Operation
		Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Seq2				= iCurColumnPos(1)
			C_WICd2				= iCurColumnPos(2)
			C_WICdPopup2		= iCurColumnPos(3)
			C_WIDesc2			= iCurColumnPos(4)
			C_ValidStartDt2		= iCurColumnPos(5)
			C_ValidEndDt2		= iCurColumnPos(6)
			C_PlantCd2			= iCurColumnPos(7)
			C_ItemCd2			= iCurColumnPos(8)
			C_RoutingNo2		= iCurColumnPos(9)
			C_OprNo2			= iCurColumnPos(10)
			C_HdnSeq2			= iCurColumnPos(11)

		' Grid 3(vspdData3) - Hidden
		Case "C"
            ggoSpread.Source = frm1.vspdData3
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Seq3				= iCurColumnPos(1)
			C_WICd3				= iCurColumnPos(2)
			C_WICdPopup3		= iCurColumnPos(3)
			C_WIDesc3			= iCurColumnPos(4)
			C_ValidStartDt3		= iCurColumnPos(5)
			C_ValidEndDt3		= iCurColumnPos(6)
			C_PlantCd3			= iCurColumnPos(7)
			C_ItemCd3			= iCurColumnPos(8)
			C_RoutingNo3		= iCurColumnPos(9)
			C_OprNo3			= iCurColumnPos(10)
			C_HdnSeq3			= iCurColumnPos(11)
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
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.id)
    If gActiveSpdSheet.id = "B" Then
		ggoSpread.Source = frm1.vspdData3
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet("C")
		Call ggoSpread.ReOrderingSpreadData()
		lgOldRow = 0
		Call vspdData1_Click(frm1.vspdData1.ActiveCol, frm1.vspdData1.ActiveRow)
	Else
		Call ggoSpread.ReOrderingSpreadData()
	End If
	
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
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
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim strCode
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	strCode = frm1.txtItemCd.value
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode			' Item Code
	arrParam(2) = "12!MO"			' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field��(0)
	arrField(1) = 2 '"ITEM_NM"			' Field��(1)
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenRoutingNo()  -------------------------------------------------
'	Name : OpenRoutingNo()
'	Description : Routing No PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRoutingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtRoutingNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
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
	
	IsOpenPop = True

	arrParam(0) = "����� �˾�"	
	arrParam(1) = "P_ROUTING_HEADER"				
	arrParam(2) = Trim(frm1.txtRoutingNo.Value)
	arrParam(3) = ""
	arrParam(4) =  "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " And ITEM_CD = " & FilterVar(frm1.txtItemCd.value, "''", "S")
	arrParam(5) = "�����"			

    arrField(0) = "ROUT_NO"	
    arrField(1) = "DESCRIPTION"	
    arrField(2) = "BOM_NO"
    arrField(3) = "MAJOR_FLG"

    arrHeader(0) = "�����"		
    arrHeader(1) = "����ø�"		
    arrHeader(2) = "BOM Type"
    arrHeader(3) = "�ֶ����"
    
    	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRoutingNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutingNo.focus
	
End Function

'------------------------------------------  OpenWcCd()  -------------------------------------------------
'	Name : OpenWcCd()
'	Description : Work center Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWcCd()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	Dim str
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	str = frm1.txtWcCd.value

	arrParam(0) = "�۾����˾�"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				  " AND VALID_TO_DT >=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & ""
	arrParam(5) = "�۾���"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "INSIDE_FLG"
    arrField(3) = "WC_MGR"	
    
    arrHeader(0) = "�۾���"		
    arrHeader(1) = "�۾����"		
    arrHeader(2) = "�۾���Ÿ��"		
    arrHeader(3) = "�۾�������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWcCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWcCd.focus
	
End Function

'------------------------------------------  OpenOprNo()  -------------------------------------------------
'	Name : OpenOprNo()
'	Description : Opr No. Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	Dim str
	
	If frm1.txtPlantCd.value = "" Then
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
	
	If frm1.txtRoutingNo.value = "" Then
		Call DisplayMsgBox("971012", "X", "�����", "X")
		frm1.txtRoutingNo.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	str = frm1.txtOprNo.value
	
	arrParam(0) = "�����˾�"	
	arrParam(1) = "P_ROUTING_DETAIL"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				  " AND VALID_TO_DT >=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & "" &_
				  " AND ITEM_CD = " & FilterVar(frm1.txtItemCd.value, "''", "S") & _
				  " AND ROUT_NO = " & FilterVar(frm1.txtRoutingNo.value, "''", "S")
	arrParam(5) = "����"			
	
    arrField(0) = "OPR_NO"
    arrField(1) = "JOB_CD"
    arrField(2) = "MILESTONE_FLG"
    
    arrHeader(0) = "����"
    arrHeader(1) = "�۾��ڵ�"
    arrHeader(2) = "Milestone"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetOprNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprNo.focus
	
End Function

'------------------------------------------  OpenWICdPopup()  -------------------------------------------------
'	Name : OpenWICD()
'	Description : Manufacturing Instruction Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWICD(ByVal str, ByVal Row)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "�����۾��˾�"	
	arrParam(1) = "P_MFG_INSTRUCTION_DETAIL"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "VALID_END_DT >=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & "" & _
				    " AND VALID_START_DT <= " &  " " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & ""
				  			
	arrParam(5) = "�����۾�"			
	
    arrField(0) = "MFG_INSTRUCTION_DTL_CD"	
    arrField(1) = "MFG_INSTRUCTION_DTL_DESC"	
    arrField(2) = "DD" & parent.gColSep & "VALID_START_DT"
    arrField(3) = "DD" & parent.gColSep & "VALID_END_DT"	
    
    arrHeader(0) = "�����۾�"		
    arrHeader(1) = "�����۾�����"		
    arrHeader(2) = "��ȿ������"		
    arrHeader(3) = "��ȿ������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWICD(arrRet, Row)
	End If
	
	Call SetActiveCell(frm1.vspdData2,C_WICd2,frm1.vspdData2.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenStdWISet()  -------------------------------------------------
'	Name : OpenStdWISet()
'	Description : Standartd Manufacturing Instruction Copy Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenStdWISet()
	Dim arrRet
	
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "ǥ���۾����ú���"						' �˾� ��Ī 
	arrParam(1) = "P_MFG_INSTRUCTION_HEADER"						' TABLE ��Ī 
	arrParam(2) = ""										' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = "VALID_TO_DT >=  " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & "" & _
				    " AND VALID_FROM_DT <= " &  " " & FilterVar(UNIConvDate(frm1.txtStdDt.Text), "''", "S") & ""										' Where Condition
	arrParam(5) = "ǥ���۾�����"							' TextBox ��Ī 
	
    arrField(0) = "MFG_INSTRUCTION_CD"									' Field��(0)
    arrField(1) = "MFG_INSTRUCTION_NM"								' Field��(0)
    
    arrHeader(0) = "ǥ���۾�����"							' Header��(0)
    arrHeader(1) = "ǥ���۾����ø�"							' Header��(0)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call  SetStdWISet(arrRet)
	End If	
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(Byval arrRet)
	frm1.txtItemCd.value= arrRet(0)
	frm1.txtItemNm.value= arrRet(1)	
End Function

'------------------------------------------  SetRouting()  --------------------------------------------------
'	Name : SetRoutingNo()
'	Description : Routing No Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRoutingNo(byval arrRet)
	frm1.txtRoutingNo.Value    = arrRet(0)
	frm1.txtRoutingNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetWcCd()  --------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWcCd(Byval arrRet)
	frm1.txtWcCd.value	= arrRet(0)
	frm1.txtWcNm.value	= arrRet(1)
End Function

'------------------------------------------  SetWcCd()  --------------------------------------------------
'	Name : SetOprNo()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOprNo(Byval arrRet)
	frm1.txtOprNo.value	= arrRet(0)
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetWICD()
'	Description : Manufacturing Instruction Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWICD(Byval arrRet, Byval Row)
	With frm1.vspdData2
		.Row = Row
		.Col = 0
'		.Text = ggoSpread.UpdateFlag
		.Col = C_WICd2
		.Text = arrRet(0)
		 .Col = C_WIDesc2
		.Text = arrRet(1)
		.Col = C_ValidStartDt2
		.Text = arrRet(2)
		.Col = C_ValidEndDt2
		.Text = arrRet(3)
		
		ggoSpread.Source = frm1.vspdData2
		.Col  = C_Seq2
		If .Text <> "" Then
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.UpdateRow Row
			CopyToHSheet Row
		End If

	End With
End Function

'------------------------------------------  SetStdWISet()  --------------------------------------------------
'	Name : SetStdWISet()
'	Description : Standard Manufacturing Instruction Set Copy Reference Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetStdWISet(byval arrRet)
		
	Dim strVal
    
    LayerShowHide(1)
		
    strVal = BIZ_PGM_COPY_ID & "?txtMode=" & parent.UID_M0001							'��: 
    strVal = strVal & "&txtStdWISet=" & arrRet(0)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtValidDt=" & Trim(frm1.hStdDt.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtRoutNo=" & Trim(frm1.hRoutNo.value)		'��: ��ȸ ���� ����Ÿ 
    frm1.vspdData1.Col = C_OprNo
    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	strVal = strVal & "&txtOprNo=" & Trim(frm1.vspdData1.Text)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData2.MaxRows
	 
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

End Function

'------------------------------------------  SetStdWISetOK()  --------------------------------------------------
'	Name : SetStdWISetOK()
'	Description : Standard Manufacturing Instruction Set Copy Reference Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetStdWISetOK(ByVal LngMaxRow)
	Call SetActiveCell(frm1.vspdData2,1,LngMaxRow,"M","X","X")
	Set gActiveElement = document.activeElement
End Function


'-------------------------------------  LookUpWI()  -----------------------------------------
'	Name : LookUp WI()
'	Description : LookUp Manufacturing Instruction
'--------------------------------------------------------------------------------------------------------- 

Function LookUpWI(Byval StrWICd, Byval Row)
    
	Dim strVal
	
    Call LayerShowHide(1)
    
    strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & parent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtWICd=" & Trim(strWICd)	'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtStdDt=" & Trim(frm1.hStdDt.value)         '��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtRow=" & Row								'��: ��ȸ ���� ����Ÿ 

    Call RunMyBizASP(MyBizASP, strVal)								'��: �����Ͻ� ASP �� ���� 
	
End Function

Function LookUpWIFail(ByRef Row)

Dim	strOprNo
Dim	strSeq

    With frm1.vspddata2
		.Row = Row
		.Col = C_WICd2
		.text = ""
		.Col = C_WIDesc2
		.text = ""
		.Col = C_ValidStartDt2
		.text = ""
		.Col = C_ValidEndDt2
		.text = ""
		.Col = C_OprNo2
		strOprNo = .text
		.Col = C_Seq2
		strSeq = .text
		Call DeleteHSheet(strOprNo, strSeq)
		
	End With
	If lgBlnFlgSaveValue =True	Then
		lgBlnFlgSaveValue = False
	End If
		
	If lgBlnFlgQryValue =True	Then
		lgBlnFlgQryValue = False
		lgBlnFlgLookupValue = True
		frm1.vspdData2.MaxRows = 0

		Call SetToolbar("11001111001111")										'��: ��ư ���� ���� 
	
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Function
			End If
	End If
	
	
End Function

Function LookUpWISuccess(ByRef strWICd, ByRef strWIDesc, ByRef strStartDt,ByRef strEndDt, ByRef Row ) 
		
	With frm1.vspdData2	
	
		.Row = Row
		.Col = C_WICd2
		.Text = UCase(strWICd)
		.Col = C_WIDesc2
		.Text = UCase(strWIDesc)
		.Col = C_ValidStartDt2
		.Text = strStartDt
		.Col = C_ValidEndDt2
		.Text = strEndDt
		
		.Col  = C_Seq2
		If .Text <> "" Then
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.UpdateRow Row
			CopyToHSheet Row
		End If
	End With
	
	If lgBlnFlgSaveValue =True	Then
		lgBlnFlgSaveValue = False
		lgBlnFlgLookupValue = True
		Call MainSave()
	End If
	
	
	If lgBlnFlgQryValue =True	Then
		lgBlnFlgQryValue = False
		lgBlnFlgLookupValue = True
		
		frm1.vspdData2.MaxRows = 0

		Call SetToolbar("11001111001111")										'��: ��ư ���� ���� 
	
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If
	
	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet("*")                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
    
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

'------------------------------------------  txtStdDt_KeyDown ----------------------------------------
'	Name : txtStdDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtStdDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtStdDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStdDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStdDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtStdDt.Focus
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Dim IntRetCD
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData1
    Call SetPopupMenuItemInf("0000110111")
	
    If frm1.vspdData1.MaxRows <= 0 Or Col < 1 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then                                                    'If there is no data.
       lgOldRow = 0
       Row = frm1.vspdData1.ActiveRow
   	End If

	If lgOldRow <> Row Then
		
		frm1.vspdData1.Col = C_OprNo
		frm1.vspdData1.Row = Row
		
		lgOldRow = Row

	    If CheckRunningBizProcess = True And lgBlnFlgLookupValue = False Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
		    lgBlnFlgQryValue = True
		    Exit Sub
		End If
		lgBlnFlgLookupValue = False
		
		frm1.vspdData2.MaxRows = 0

		Call SetToolbar("11001111001111")										'��: ��ư ���� ����	
	
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If	
		
	End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	Dim IntRetCD
    gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1101110111")
	Else
		Call SetPopupMenuItemInf("0000110111")
	End If

    If frm1.vspdData2.MaxRows <= 0 Or Col < 1 Or Row <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData1

		If Row = NewRow Then
		    Exit Sub
		End If

		If NewRow <= 0 Or NewCol < 0 Then
			Exit Sub
		End If

	
		Call SetToolbar("11001011000111")										'��: ��ư ���� ���� 
	
    End With

End Sub

'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)

	Dim strSeq, strWICd
	Dim strHndSeq, strHndOprNo
	Dim i
	Dim strReqDt, strEndDt
	Dim	DblRqrdQty, DblIssuedQty
	Dim lNewRow, lOldRow

	lOldRow = frm1.vspdData1.ActiveRow
					
	With frm1.vspdData2

		Select Case Col

		    Case C_Seq2

				.Row = Row
				.Col = C_Seq2
				strSeq = .Text
				
				If strSeq = "" Then Exit Sub
				
				For i = 1 To .MaxRows
					If i <> Row Then
						.Row = i
						.Col = C_Seq2
						If UCase(Trim(.Text)) = UCase(Trim(strSeq)) Then
							Call DisplayMsgBox("181416", "X", UCase(Trim(strSeq)), "X")
							.Row = Row
							.Text = ""
							Exit Sub
						End If
					End If						
				Next
				
				.Row = Row
				.Col = C_OprNo2
				strHndOprNo = .Text 				
				.Col = C_HdnSeq2
				strHndSeq = .Text
			
				If strHndSeq <> "" Then
					Call DeleteHSheet(strHndOprNo, strHndSeq)
				End If
				
				.Row = Row
				.Col = C_HdnSeq2
				.Text = strSeq

				.Row = Row
				.Col = C_WICd2

				If .Text <> "" Then
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.UpdateRow Row
					CopyToHSheet Row
				End If
				

		    Case C_WICd2
				
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				.Row = Row
				.Col = C_WICd2
				strWICd = .Text

				If .Text <> "" Then	
					Call LookUpWI(strWICd, Row)
	
				End If
			
		End Select

	End With

End Sub

'=======================================================================================================
'   Function Name : FixHiddenRow
'   Function Desc : 
'=======================================================================================================
Function FixHiddenRow(Byval strOprNo, Byval strItemCd, Byval Col, Byval strValue)

Dim strHndOprNo, strHndItemCd
Dim lRows

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows
        
            .vspdData3.Row = lRows
            .vspdData3.Col = C_OprNo3
            strHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_CompntCd3
            strHndItemCd = .vspdData3.Text

            If Trim(strHndOprNo) = Trim(strOprNo) And Trim(strHndItemCd) = Trim(strItemCd) Then
				.vspdData3.Col = Col
				.vspdData3.Text = strValue
				ggoSpread.Source = frm1.vspdData3
				ggoSpread.UpdateRow lRows
				Exit Function
            End If    
        Next
        
    End With        
    
End Function

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

Dim strCode
Dim strName

    With frm1.vspdData2
    
		ggoSpread.Source = frm1.vspdData2
		If Row < 1 Then Exit Sub

		Select Case Col

		    Case C_WICdPopup2
				.Col = C_WICd2
				.Row = Row
				strCode = .Text
				Call OpenWICd(strCode, Row)
				
				Call SetActiveCell(frm1.vspdData2,C_WICd2,Row,"M","X","X")
				Set gActiveElement = document.activeElement
	     
		End Select

	End With

End Sub


'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If  
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then

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

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False											'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
	Call ggoSpread.ClearSpreadData
    Call InitVariables
    frm1.vspdData1.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
    frm1.vspdData3.MaxRows = 0
	lgFlgQueryCnt = 0

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		 Call RestoreToolBar()
		 Exit Function												'��: Query db data		 
    End If   
    FncQuery = True												'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 

    Dim IntRetCD 
    Dim	LngRows
    
    If CheckRunningBizProcess = True And lgBlnFlgLookupValue = False Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        lgBlnFlgSaveValue = True
        Exit Function
	End If
	
    lgBlnFlgLookupValue = False
    
    FncSave = False												'��: Processing is NG
    
    Err.Clear                                                   '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
       
    ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
       
    ggoSpread.Source = frm1.vspdData3							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
 
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function													'��: Save db data
    
    FncSave = True												'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
        
	If frm1.vspdData2.MaxRows < 1 Then Exit Function	
        
    ggoSpread.Source = frm1.vspdData2
	frm1.vspdData2.ReDraw = False
	If frm1.vspdData2.ActiveRow > 0 Then
       ggoSpread.CopyRow
		Call SetSpreadColor(frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow)
		frm1.vspdData2.ReDraw = True
		frm1.vspdData2.Focus
    End If
    
    frm1.vspdData2.Col = C_Seq2
    frm1.vspdData2.Text = ""
    frm1.vspdData2.Col = C_HdnSeq2
    frm1.vspdData2.Text = ""
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()

Dim strMode
Dim	strOprNo
Dim	strSeq

	If frm1.vspdData2.MaxRows < 1 Then Exit Function	

    ggoSpread.Source = frm1.vspdData2
    Call frm1.vspdData2.GetText(0, frm1.vspdData2.ActiveRow, strMode)
    Call frm1.vspdData2.GetText(C_OprNo2, frm1.vspdData2.ActiveRow, strOprNo)
    Call frm1.vspdData2.GetText(C_Seq2, frm1.vspdData2.ActiveRow, strSeq)

	ggoSpread.EditUndo
	Call EditUndoHSheet(strOprNo, strSeq)

	If strMode = ggoSpread.UpdateFlag Then
	    Call Copy1RowFromHSheet(strOprNo, strSeq, frm1.vspdData2.ActiveRow)
	End If

End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 
    Dim iIntCnt, iIntReqRows

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If

    With frm1.vspdData2
		.ReDraw = False
		.Focus
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.InsertRow , iIntReqRows

        Call SetSpreadColor(.ActiveRow, .ActiveRow + iIntReqRows - 1)
		For iIntCnt = .ActiveRow To .ActiveRow + iIntReqRows - 1
			.Row = iIntCnt
			.Col = C_PlantCd2
			.Text = UCase(Trim(frm1.hPlantCd.value))
			.Col = C_ItemCd2
			.Text = UCase(Trim(frm1.hItemCd.value))
			.Col = C_RoutingNo2
			.Text = UCase(Trim(frm1.hRoutNo.value))
			frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
			frm1.vspdData1.Col = C_OprNo
			.Col = C_OprNo2
			.Text = UCase(Trim(frm1.vspdData1.Text))
		Next
		.ReDraw = True
    End With

End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt, i

    With frm1
		
   
		If .vspdData2.MaxRows < 1 Then Exit Function

		Call DeleteMarkingHSheet()

    End With

	ggoSpread.Source = frm1.vspdData2
    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows

	CopyToHSheet frm1.vspdData2.ActiveRow

End Function

'=======================================================================================================
'   Function Name : DeleteMarkingHSheet
'   Function Desc : DeleteMark the Row Which keys match with vapdData's Key and vspdData2's Key
'=======================================================================================================
Function DeleteMarkingHSheet()

	Dim lRow, lRows
	
	Dim strInspItemCd
	Dim strInspSeries
	Dim strSampleNo
	Dim lngRow2
	Dim strHndOprNo, strOprNo, strHndSeq, strSeq	
	
	DeleteMarkingHSheet = False
		
	For lngRow2 = frm1.vspdData2.SelBlockRow To frm1.vspdData2.SelBlockRow2
	
        For lRows = 1 To frm1.vspdData3.MaxRows
            frm1.vspdData3.Row = lRows
            frm1.vspdData3.Col = C_OprNo3
            strHndOprNo = frm1.vspdData3.Text
            frm1.vspdData3.Col = C_Seq3
            strHndSeq = frm1.vspdData3.Text
            frm1.vspdData2.Row = lngRow2
            frm1.vspdData2.Col = C_OprNo2
            strOprNo = frm1.vspdData2.Text
            frm1.vspdData2.Col = C_Seq2
            strSeq = frm1.vspdData2.Text
            If strHndOprNo = strOprNo And strHndSeq = strSeq Then
				lRow = lRows
				Exit For
            End If    
		Next
	
		If lRow > 0 Then
			With frm1
    			ggoSpread.Source = .vspdData3
		 		.vspdData3.Col = 0
				.vspdData3.Text = ggoSpread.DeleteFlag
			End With
		End If
	Next
	
	DeleteMarkingHSheet = True
	
End Function    

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()                                                   '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												<%'��: ȭ�� ���� %>
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                                    <%'��: Protect system from crashing%>
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData2							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")	'��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim strVal
    
    DbQuery = False
    
    lgFlgQueryCnt = lgFlgQueryCnt + 1
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing

    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode	    
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" &  Trim(.hItemCd.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtRoutNo=" &  Trim(.hRoutNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtOprNo=" &  Trim(.hOprNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" &  Trim(.hWcCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtStdDt=" &  Trim(.hStdDt.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	Else
		strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" &  Trim(.txtItemCd.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtRoutNo=" &  Trim(.txtRoutingNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtOprNo=" &  Trim(.txtOprNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtWcCd=" &  Trim(.txtWcCd.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtStdDt=" &  Trim(.txtStdDt.Text)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

	Call SetToolbar("11001111001111")										'��: ��ư ���� ���� 
	
	frm1.vspdData1.Col = C_OprNo
	frm1.vspdData1.Row = 1

	lgOldRow = 1

	If lgFlgQueryCnt = 1 Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
			Set gActiveElement = document.activeElement
			If DbDtlQuery = False Then
			
			End If
		End If
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOprCd
    
	boolExist = False
    With frm1

	    Call .vspdData1.GetText(C_OprNo, .vspdData1.ActiveRow, strOprCd)
    
	    If CopyFromHSheet(strOprCd) = True Then
           Exit Function
        End If

		DbDtlQuery = False   
    
		.vspdData1.Row = .vspdData1.ActiveRow

		Call LayerShowHide(1)       
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd=" &  Trim(.hItemCd.Value)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtRoutNo=" &  Trim(.hRoutNo.Value)		'��: ��ȸ ���� ����Ÿ 
			.vspdData1.Col = C_OprNo
			strVal = strVal & "&txtOprNo=" & Trim(.vspdData1.Text)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtStdDt=" &  Trim(.hStdDt.Value)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtMaxRows2=" & .vspdData2.MaxRows
			strVal = strVal & "&txtMaxRows3=" & .vspdData3.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd=" &  Trim(.txtItemCd.Value)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtRoutNo=" &  Trim(.txtRoutingNo.Value)		'��: ��ȸ ���� ����Ÿ 
			.vspdData1.Col = C_OprNo
			strVal = strVal & "&txtOprNo=" & Trim(.vspdData1.Text)
			strVal = strVal & "&txtStdDt=" &  Trim(.txtStdDt.Text)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtMaxRows2=0"
			strVal = strVal & "&txtMaxRows3=0"
			
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    End With

    DbDtlQuery = True

End Function

'========================================================================================
' Function Name : DbDtlQueryOk
' Function Desc : DbDtlQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbDtlQueryOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 

	Dim LngRow

    '-----------------------
    'Reset variables area
    '-----------------------
	frm1.vspdData2.ReDraw = False
	
	ggoSpread.Source = frm1.vspdData2

   
	lgAfterQryFlg = True

	frm1.vspdData2.ReDraw = True

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData(ByVal Row)

Dim strOprNo,strSeq
Dim strHndOprNo, strHndSeq
Dim lRows

    FindData = 0

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows
        
            .vspdData3.Row = lRows
            .vspdData3.Col = C_OprNo3
            strHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_Seq3
            strHndSeq = .vspdData3.Text
            .vspdData2.Row = Row
            .vspdData2.Col = C_OprNo2
            strOprNo = .vspdData2.Text
            .vspdData2.Col = C_HdnSeq2
            strSeq = .vspdData2.Text
            
            If Trim(strHndOprNo) = Trim(strOprNo) And Trim(strHndSeq) = Trim(strSeq) Then
				FindData = lRows
				Exit Function
            End If    
        Next
        
    End With        
    
End Function


'=======================================================================================================
'   Function Name : CopyFromHSheet
'   Function Desc : 
'=======================================================================================================
Function CopyFromHSheet(ByVal strOprNo)

Dim lngRows
Dim boolExist
Dim iCols
Dim strHdnOprNo
Dim strStatus
Dim iCurColumnPos2

    boolExist = False
    
    CopyFromHSheet = boolExist

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos2)
    
    With frm1

        Call SortHSheet()
			
		'------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
			Call .vspdData3.GetText(C_OprNo3, lngRows, strHdnOprNo)
            If  strOprNo = strHdnOprNo Then             
                boolExist = True
                Exit For
            End If    
        Next

		'------------------------------------
		' Show Data
		'------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            frm1.vspdData2.Redraw = False
            
            While lngRows <= .vspdData3.MaxRows

	            .vspdData3.Row = lngRows
                
                .vspdData3.Col = C_OprNo3
				strHdnOprNo = .vspdData3.Text
                
                If strOprNo <> strHdnOprNo Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
					If strOprNo = strHdnOprNo Then
						.vspdData2.MaxRows = .vspdData2.MaxRows + 1
						.vspdData2.Row = .vspdData2.MaxRows
						.vspdData2.Col = 0
						.vspdData3.Col = 0
						.vspdData2.Text = .vspdData3.Text
						
						For iCols = 1 To .vspdData3.MaxCols
						    .vspdData2.Col = iCurColumnPos2(iCols)
						    .vspdData3.Col = iCols
						    .vspdData2.Text = .vspdData3.Text
						Next
						
						.vspdData3.Col = 0
						If .vspdData3.Text = ggoSpread.InsertFlag Then 
							ggoSpread.Source = frm1.vspdData2
							ggoSpread.SpreadUnLock  C_Seq2,	lngRows, C_Seq2, lngRows
							ggoSpread.SSSetRequired C_Seq2,	lngRows, lngRows
						End If
			
					End If
                End If   
                
                lngRows = lngRows + 1
                
            Wend
            frm1.vspdData2.Redraw = True

        End If
            
    End With        
    
    CopyFromHSheet = boolExist
   
End Function


'=======================================================================================================
'   Function Name : Copy1RowFromHSheet
'   Function Desc : 
'=======================================================================================================
Function Copy1RowFromHSheet(ByVal strOprNo, ByVal strSeq, ByVal pvRow)

Dim lngRows
Dim iCols
Dim strHdnOprNo
Dim strHndSeq
Dim iCurColumnPos2

	On Error Resume Next
	Err.Clear

    Copy1RowFromHSheet = False
    
    With frm1
        For lngRows = 1 To .vspdData3.MaxRows
			Call .vspdData3.GetText(C_OprNo3, lngRows, strHdnOprNo)
			Call .vspdData3.GetText(C_HdnSeq3, lngRows, strHndSeq)

            If strOprNo = strHdnOprNo And strSeq = strHndSeq Then
                ggoSpread.Source = .vspdData3
                ggoSpread.EditUndo lngRows
                ggoSpread.Source = .vspdData2
				Call ggoSpread.GetSpreadColumnPos(iCurColumnPos2)

				.vspdData3.Row = lngRows
				.vspdData2.Row = pvRow
				For iCols = 1 To .vspdData3.MaxCols
				    .vspdData2.Col = iCurColumnPos2(iCols)
				    .vspdData3.Col = iCols
				    .vspdData2.Text = .vspdData3.Text
				Next

				Copy1RowFromHSheet = True
                Exit For
            End If    
        Next
    End With

End Function

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 
'=======================================================================================================
Sub CopyToHSheet(ByVal Row)
	Dim lRow
	Dim iCols
	Dim iCurColumnPos2

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos2)

	With frm1 
        
	    lRow = FindData(Row)

	    If lRow > 0 Then
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            For iCols = 1 To 6 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCurColumnPos2(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next
            
			.vspdData2.Col = C_Seq2
			.vspdData3.Col = C_HdnSeq3
			.vspdData3.Text = .vspdData2.Text
            
        Else
			.vspdData3.MaxRows = .vspdData3.MaxRows + 1
            .vspdData3.Row = .vspdData3.MaxRows
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
       
            For iCols = 1 To .vspdData2.MaxCols 'vspdData2 �� ����Ÿ�� �����Ѵ�.
                .vspdData2.Col = iCurColumnPos2(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next

			.vspdData2.Col = C_Seq2
			.vspdData3.Col = C_HdnSeq3
			.vspdData3.Text = .vspdData2.Text
        
        End If

	End With
	
End Sub

'=======================================================================================================
'   Function Name : EditUndoHSheet
'   Function Desc : 
'=======================================================================================================
Function EditUndoHSheet(ByVal strOprNo, Byval strSeq)

Dim lngRows
Dim StrHndOprNo, strHndSeq
 
    EditUndoHSheet = False
    
    With frm1
    
        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
			Call .vspdData3.GetText(C_OprNo3, lngRows, StrHndOprNo)
			Call .vspdData3.GetText(C_HdnSeq3, lngRows, strHndSeq)

            If strOprNo = StrHndOprNo And strSeq = strHndSeq Then
				ggoSpread.Source = .vspdData3
				ggoSpread.EditUndo lngRows
				EditUndoHSheet = True
                Exit For
            End If    
        Next
        
    End With

End Function   

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal strOprNo, Byval strSeq)

Dim boolExist
Dim lngRows
Dim StrHndOprNo, strHndSeq
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_OprNo3
			StrHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_HdnSeq3
			strHndSeq = .vspdData3.Text

            If strOprNo = StrHndOprNo and strSeq = strHndSeq Then
                boolExist = True
                Exit For
            End If    
        Next
       
        '------------------------------------
        ' Data Delete
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
				.vspdData3.Col = C_OprNo3
				StrHndOprNo = .vspdData3.Text
				.vspdData3.Col = C_HdnSeq3
				strHndSeq = .vspdData3.Text
                
                If (strOprNo <> StrHndOprNo) or (strSeq <> strHndSeq) Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If   

            Wend
            
            ggoSpread.Source = frm1.vspdData2
            
            frm1.vspdData2.Row = lgCurrRow
            frm1.vspdData2.Col = frm1.vspdData2.MaxCols
            ggoSpread.Source = frm1.vspdData2

            frm1.vspdData2.Redraw = True

        End If

    End With

    DeleteHSheet = True
End Function    

'======================================================================================================
' Function Name : SortHSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortHSheet()
    
    With frm1
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0 'SS_SORT_BY_ROW
        
        .vspdData3.SortKey(1) = C_OprNo3		' Operation No
        .vspdData3.SortKey(2) = C_Seq3		    ' Sequence No
        
        .vspdData3.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25 'SS_ACTION_SORT
        .vspdData3.BlockMode = False
    End With        
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

   Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	Dim iColSep
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt
	
    DbSave = False                                                          '��: Processing is NG
	     
    LayerShowHide(1)
		
	With frm1
		
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
    iValCnt = 0 : iDelCnt = 0
    
    '-----------------------
    'Data manipulate area
    '-----------------------
        
    For lRow = 1 To .vspdData3.MaxRows
    
        .vspdData3.Row = lRow
        .vspdData3.Col = 0

        Select Case .vspdData3.Text
				
			   Case ggoSpread.InsertFlag                                      '��: Update
														   strVal = ""
                                                           strVal = strVal & "C" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData3.Col = C_PlantCd3	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_ItemCd3	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_RoutingNo3	       : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_OprNo3	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_Seq3		           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_WICd3 	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gRowSep
                    ReDim Preserve TmpBufferVal(iValCnt)
                    TmpBufferVal(iValCnt) = strVal
                    iValCnt = iValCnt + 1
                    
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '��: Update
														   strVal = ""
                                                           strVal = strVal & "U" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData3.Col = C_PlantCd3	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_ItemCd3	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_RoutingNo3	       : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_OprNo3	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_Seq3		           : strVal = strVal & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_WICd3 	           : strVal = strVal & Trim(.vspdData3.Text) & parent.gRowSep
                    
                    ReDim Preserve TmpBufferVal(iValCnt)
                    TmpBufferVal(iValCnt) = strVal
                    iValCnt = iValCnt + 1
                    
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.DeleteFlag                                      '��: Delete
														   strDel = ""
                                                           strDel = strDel & "D" & parent.gColSep
                                                           strDel = strDel & lRow & parent.gColSep
                    .vspdData3.Col = C_PlantCd3	           : strDel = strDel & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_ItemCd3	           : strDel = strDel & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_RoutingNo3	       : strDel = strDel & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_OprNo3	           : strDel = strDel & Trim(.vspdData3.Text) & parent.gColSep
                    .vspdData3.Col = C_Seq3		           : strDel = strDel & Trim(.vspdData3.Text) & parent.gRowSep
                    
                    ReDim Preserve TmpBufferDel(iDelCnt)
                    TmpBufferDel(iDelCnt) = strDel
                    iDelCnt = iDelCnt + 1
                    
                    lGrpCnt = lGrpCnt + 1
                   
				
	    End Select
                
    Next
	
	iTotalStrDel = Join(TmpBufferDel, "")
	iTotalStrVal = Join(TmpBufferVal, "")
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = iTotalStrDel & iTotalStrVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           ' ��: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0

	ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0
	ggoSpread.source = frm1.vspddata3
    frm1.vspdData3.MaxRows = 0
	
	Call DbDtlQuery
	
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!-- '#########################################################################################################    
            					6. Tag�� 
'######################################################################################################## -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�۾����õ��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenStdWISet()">ǥ���۾�����</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="12xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="24"></TD></TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtRoutingNo" SIZE=10 MAXLENGTH=7 tag="12xxxU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRoutingNo()">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="11xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()"></TD></TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWcCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p1208ma1_I518603992_txtStdDt.js'></script>
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
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p1208ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p1208ma1_B_vspdData2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hOprNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hStdDt" tag="24">
<script language =javascript src='./js/p1208ma1_C_vspdData3.js'></script>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
