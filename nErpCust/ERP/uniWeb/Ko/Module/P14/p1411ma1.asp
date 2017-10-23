<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1411ma1.asp
'*  4. Program Name         : BOM�����̷� ��ȸ 
'*  5. Program Desc         :
'*  6. Component List        : 
'*  7. Modified date(First) : 2003/03/08
'*  8. Modified date(Last)  : 2003/03/08
'*  9. Modifier (First)     : NamkyuHo
'* 10. Modifier (Last)      : Park Kye Jin (Reference Popup Added) (2003.04.10)
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1411mb11.asp"											'��: �����Ͻ� ���� ASP�� 

Dim	C_Itemcd
Dim	C_ItemNm
Dim	C_Spec
Dim	C_ActionFlg
Dim	C_InsertDt
Dim	C_InsertUserId
Dim	C_ChangeSeq
Dim C_ChildSeq
Dim	C_ChildItemCd
Dim	C_ChildItemNm
Dim	C_ChildItemSpec
Dim	C_Acct
Dim	C_ProcurType
Dim	C_ChildItemQty
Dim	C_ChildUnit
Dim	C_PrntItemQty
Dim	C_PrntUnit
Dim	C_SafetyLT
Dim	C_LossRate
Dim	C_SupplyFlg
Dim	C_ValidFromDt
Dim	C_ValidToDt
Dim	C_ECNNo
Dim	C_ECNDescription
Dim	C_ECNReasonCd
Dim	C_Remark
Dim C_ChangedField
Dim	C_Seq
Dim C_InsertDtHD

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
       
Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Itemcd			= 1	
	C_ItemNm			= 2
	C_Spec				= 3	
	C_ActionFlg			= 4
	C_InsertDt			= 5
	C_InsertUserId		= 6
	C_ChangeSeq			= 7
	C_ChildSeq			= 8
	C_ChildItemCd		= 9
	C_ChildItemNm		= 10
	C_ChildItemSpec		= 11
	C_Acct				= 12
	C_ProcurType		= 13
	C_ChildItemQty		= 14
	C_ChildUnit			= 15
	C_PrntItemQty		= 16
	C_PrntUnit			= 17
	C_SafetyLT			= 18
	C_LossRate			= 19
	C_SupplyFlg			= 20
	C_ValidFromDt		= 21
	C_ValidToDt			= 22
	C_ECNNo				= 23
	C_ECNDescription	= 24
	C_ECNReasonCd		= 25
	C_Remark			= 26
	C_ChangedField		= 27
	C_InsertDtHD		= 28
	C_Seq				= 29
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKeyIndex = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                                       '��: initializes sort direction
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtChgFromDt.Text = UNIDateAdd("D",-10, StartDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
	frm1.txtChgToDt.Text = StartDate
	frm1.txtBomNo.value = "1"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "MA")%>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030308",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_Seq												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd, 			"ǰ��",		20,,, 18, 2
		ggoSpread.SSSetEdit		C_ItemNm,			"ǰ���",	30
		ggoSpread.SSSetEdit		C_Spec,				"�԰�",		30
		ggoSpread.SSSetEdit		C_ActionFlg,		"���躯�汸��", 12
		ggoSpread.SSSetEdit		C_InsertDt,			"���躯����",	20
		ggoSpread.SSSetEdit		C_InsertDtHD,		"���躯����",	20
		ggoSpread.SSSetEdit		C_InsertUserId,		"���躯����",	13
		ggoSpread.SSSetFloat	C_ChangeSeq,		"���躯�����", 12, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetFloat	C_ChildSeq,			"��ǰ�����",	10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetEdit		C_ChildItemcd,		"��ǰ��",	20,,, 18, 2
		ggoSpread.SSSetEdit		C_ChildItemNm,		"��ǰ���", 30
		ggoSpread.SSSetEdit		C_ChildItemSpec,	"��ǰ��԰�", 30
		ggoSpread.SSSetEdit		C_Acct,				"ǰ�����", 10
		ggoSpread.SSSetEdit		C_ProcurType,		"���ޱ���", 12
		ggoSpread.SSSetFloat	C_ChildItemQty,		"��ǰ����ؼ�", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
		ggoSpread.SSSetEdit		C_ChildUnit,		"����",		6,,, 3, 2
		ggoSpread.SSSetFloat	C_PrntItemQty,		"��ǰ����ؼ�", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
		ggoSpread.SSSetEdit		C_PrntUnit,			"����",		6,,, 3, 2
		ggoSpread.SSSetFloat	C_SafetyLT,			"����L/T",	10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetFloat	C_LossRate,			"Loss��",	10, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z"
		ggoSpread.SSSetEdit		C_SupplyFlg,		"�����󱸺�", 8
		ggoSpread.SSSetDate		C_ValidFromDt,		"������",	11, 2, parent.gDateFormat
		ggoSpread.SSSetDate		C_ValidToDt,		"������",	11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ECNNo,			"���躯���ȣ", 18
		ggoSpread.SSSetEdit		C_ECNDescription,	"���躯�泻��", 20
		ggoSpread.SSSetEdit		C_ECNReasonCd,		"���躯��ٰ�", 15
		ggoSpread.SSSetEdit		C_Remark,			"���", 30
		ggoSpread.SSSetEdit		C_ChangedField,		"������ʵ�", 20,,, 150, 2
		ggoSpread.SSSetEdit		C_Seq,				"����", 5
		
		ggoSpread.SSSetSplit2(1)											'frozen ��� �߰� 
		
		Call ggoSpread.MakePairsColumn(C_ChildItemQty, C_ChildUnit)
		Call ggoSpread.MakePairsColumn(C_PrntItemQty, C_PrntUnit)
		
		Call ggoSpread.SSSetColHidden(C_InsertDtHD, C_InsertDtHD, True)
		Call ggoSpread.SSSetColHidden(C_ChangedField, C_ChangedField, True)
		Call ggoSpread.SSSetColHidden(C_Seq, C_Seq, True)
    
		.ReDraw = True

		Call SetSpreadLock 

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
			
			C_Itemcd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_ActionFlg		= iCurColumnPos(4)
			C_InsertDt		= iCurColumnPos(5)
			C_InsertUserId	= iCurColumnPos(6)
			C_ChangeSeq		= iCurColumnPos(7)
			C_ChildSeq		= iCurColumnPos(8)
			C_ChildItemCd	= iCurColumnPos(9)
			C_ChildItemNm	= iCurColumnPos(10)
			C_ChildItemSpec	= iCurColumnPos(11)
			C_Acct			= iCurColumnPos(12)
			C_ProcurType	= iCurColumnPos(13)
			C_ChildItemQty	= iCurColumnPos(14)
			C_ChildUnit		= iCurColumnPos(15)
			C_PrntItemQty	= iCurColumnPos(16)
			C_PrntUnit		= iCurColumnPos(17)
			C_SafetyLT		= iCurColumnPos(18)
			C_LossRate		= iCurColumnPos(19)
			C_SupplyFlg		= iCurColumnPos(20)
			C_ValidFromDt	= iCurColumnPos(21)
			C_ValidToDt		= iCurColumnPos(22)
			C_ECNNo			= iCurColumnPos(23)
			C_ECNDescription = iCurColumnPos(24)
			C_ECNReasonCd	= iCurColumnPos(25)
			C_Remark		= iCurColumnPos(26)
			C_ChangedField	= iCurColumnPos(27)
			C_InsertDtHD	= iCurColumnPos(28)
			C_Seq			= iCurColumnPos(29)
					
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
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
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
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
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

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	
	arrParam(2) = Trim(frm1.txtBomNo.value)		' Code Condition
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
	Frm1.txtBomNo.Focus
	
End Function

 '------------------------------------------  OpenECNNo()  -------------------------------------------------
 '	Name : OpenECNNo()
 '	Description : ECN PopUp
 '--------------------------------------------------------------------------------------------------------- 
 Function OpenECNNo()
 	Dim arrRet
 	Dim arrParam(4), arrField(10)
 	Dim iCalledAspName, IntRetCD
 
 	If IsOpenPop = True Then 
 		IsOpenPop = False
 		Exit Function
 	End If
 	
 	IsOpenPop = True
 	
 	arrParam(0) = frm1.txtECNNo.value   ' ECN No.
 
 	iCalledAspName = AskPRAspName("P1410PA1")
 	
 	If Trim(iCalledAspName) = "" Then
 		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
 		IsOpenPop = False
 		Exit Function
 	End If
 	
 	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
 		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
 	
 	IsOpenPop = False
 	
 	If arrRet(0) <> "" Then
 		Call SetECNNo(arrRet)
 	End If	
 	
 	Call SetFocusToDocument("M")
	Frm1.txtECNNo.Focus
 	
 End Function
 
'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"	
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
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
	Frm1.txtItemCd.Focus
	
End Function

'------------------------------------------  OpenChildItemCd()  -------------------------------------------------
'	Name : OpenChildIremCd()
'	Description : Child Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenChildItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(frm1.txtChildItemCd.value)	' Item Code
	arrParam(2) = ""					' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "CHILD_ITEM_CD"	
    arrField(1) = 2 							' Field��(1) : "CHILD_ITEM_NM"	
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetChildItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	Frm1.txtChildItemCd.Focus
	
End Function
'------------------------------------------  OpenChangeRef()  ----------------------------------------------
'	Name : OpenChangeRef()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenChangeRef()
	Dim arrRet
	Dim arrParam(26)
	Dim iCalledAspName
	Dim PrevPrntItem
	Dim PrevChildItem
	Dim PrevSeq
	Dim SaveSeq
	Dim i

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
  
	ggoSpread.Source = frm1.vspdData    
	With frm1.vspdData     
		If .MaxRows = 0 Then
			Call DisplayMsgBox("169903","X", "X", "X")    'ǰ���ڷᰡ �ʿ��մϴ� 
			Exit Function
		End If 

		If .ActiveRow = 1 Then
			.Col = C_ChangedField
			.Row = .ActiveRow
			arrParam(0) = Trim(.Text )
			.Col = C_Itemcd
			.Row = .ActiveRow
			arrParam(1) = Trim(.Text )
			.Col = C_ItemNm
			.Row = .ActiveRow
			arrParam(2) = Trim(.Text )
			.Col = C_ChildItemCd
			.Row = .ActiveRow
			arrParam(3) = Trim(.Text )
			.Col = C_ChildItemNm
			.Row = .ActiveRow
			arrParam(4) = Trim(.Text)
			.Col = C_ChildSeq
			.Row = .ActiveRow
			arrParam(5) = Trim(.Text)
			.Col = C_ActionFlg
			.Row = .ActiveRow
			arrParam(6) = Trim(.Text)
			.Col = C_ChildItemQty
			.Row = .ActiveRow
			arrParam(7) = Trim(.Text)
			.Col = C_ChildUnit
			.Row = .ActiveRow
			arrParam(8) = Trim(.Text)
			.Col = C_PrntItemQty
			.Row = .ActiveRow
			arrParam(9) = Trim(.Text)
			.Col = C_PrntUnit
			.Row = .ActiveRow
			arrParam(10) = Trim(.Text)
			.Col = C_SafetyLT
			.Row = .ActiveRow
			arrParam(11) = Trim(.Text)
			.Col = C_LossRate
			.Row = .ActiveRow
			arrParam(12) = Trim(.Text)
			.Col = C_SupplyFlg
			.Row = .ActiveRow
			arrParam(13) = Trim(.Text)
			.Col = C_ValidFromDt
			.Row = .ActiveRow
			arrParam(14) = Trim(.Text)
			.Col = C_ValidToDt
			.Row = .ActiveRow
			arrParam(15) = Trim(.Text)
			arrParam(16) = ""
			arrParam(17) = ""
			arrParam(18) = ""
			arrParam(19) = ""
			arrParam(20) = ""
			arrParam(21) = ""
			arrParam(22) = ""
			arrParam(23) = ""
			arrParam(24) = ""
			.Col = C_InsertDt
			.Row = .ActiveRow
			arrParam(25) = Trim(.Text)
			.Col = C_InsertUserId
			.Row = .ActiveRow
			arrParam(26) = Trim(.Text)
		else 
			.Col = C_Seq
			.Row = .ActiveRow
			SaveSeq = Trim(.Text )
			.Col = C_ChangedField
			.Row = .ActiveRow
			arrParam(0) = Trim(.Text )
			.Col = C_Itemcd
			.Row = .ActiveRow
			arrParam(1) = Trim(.Text )
			.Col = C_ItemNm
			.Row = .ActiveRow
			arrParam(2) = Trim(.Text )
			.Col = C_ChildItemCd
			.Row = .ActiveRow
			arrParam(3) = Trim(.Text )
			.Col = C_ChildItemNm
			.Row = .ActiveRow
			arrParam(4) = Trim(.Text)
			.Col = C_ChildSeq
			.Row = .ActiveRow
			arrParam(5) = Trim(.Text)
			.Col = C_ActionFlg
			.Row = .ActiveRow
			arrParam(6) = Trim(.Text)
			.Col = C_ChildItemQty
			.Row = .ActiveRow
			arrParam(7) = Trim(.Text)
			.Col = C_ChildUnit
			.Row = .ActiveRow
			arrParam(8) = Trim(.Text)
			.Col = C_PrntItemQty
			.Row = .ActiveRow
			arrParam(9) = Trim(.Text)
			.Col = C_PrntUnit
			.Row = .ActiveRow
			arrParam(10) = Trim(.Text)
			.Col = C_SafetyLT
			.Row = .ActiveRow
			arrParam(11) = Trim(.Text)
			.Col = C_LossRate
			.Row = .ActiveRow
			arrParam(12) = Trim(.Text)
			.Col = C_SupplyFlg
			.Row = .ActiveRow
			arrParam(13) = Trim(.Text)
			.Col = C_ValidFromDt
			.Row = .ActiveRow
			arrParam(14) = Trim(.Text)
			.Col = C_ValidToDt
			.Row = .ActiveRow
			arrParam(15) = Trim(.Text)
			.Col = C_InsertDt
			.Row = .ActiveRow
			arrParam(25) = Trim(.Text)
			.Col = C_InsertUserId
			.Row = .ActiveRow
			arrParam(26) = Trim(.Text)
		End If 

			.Col = C_Seq					' Grid Sort �� �����ʵ带 ã������ Logic KJPark
			For i = 1 to .MaxRows   
				.Row = i
				If Trim(SaveSeq - 1) = Trim(.Text ) then 
					.Col = C_Itemcd
					PrevPrntItem = Trim(.Text )

					.Col = C_ChildItemCd
					PrevChildItem = Trim(.Text )

					.Col = C_ChildSeq
					PrevSeq = Trim(.Text)

					.Col = C_ChildItemQty
					arrParam(16) = Trim(.Text)

					.Col = C_ChildUnit
					arrParam(17) = Trim(.Text)

					.Col = C_PrntItemQty
					arrParam(18) = Trim(.Text)

					.Col = C_PrntUnit
					arrParam(19) = Trim(.Text)

					.Col = C_SafetyLT
					arrParam(20) = Trim(.Text)

					.Col = C_LossRate
					arrParam(21) = Trim(.Text)

					.Col = C_SupplyFlg
					arrParam(22) = Trim(.Text)

					.Col = C_ValidFromDt
					arrParam(23) = Trim(.Text)

					.Col = C_ValidToDt
					arrParam(24) = Trim(.Text)
					Exit For
				End If
			Next 	
							
		If  arrParam(6) = "Change" and (PrevPrntItem <> arrParam(1) or PrevChildItem <> arrParam(3) or PrevSeq <>arrParam(5)) Then
			Call DisplayMsgBox("200002","X", "X", "X")    '�����ڷᰡ �����ϴ� 
			Exit Function
		End If 
		
	End With
	

	IsOpenPop = True

	iCalledAspName = AskPRAspName("P1411RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"P1411RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent , arrParam), _
	"dialogWidth=760px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")      
	
	Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
	Set gActiveElement = document.activeElement
	     
	IsOpenPop = False

End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	
	frm1.txtBOMNo.focus
	Set gActiveElement = document.activeElement		
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	frm1.txtBomNo.Value    = arrRet(0)	
	
	frm1.txtECNNo.focus
	Set gActiveElement = document.activeElement		
End Function

'------------------------------------------  SetECNNo()  --------------------------------------------------
'	Name : SetECNNo()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetECNNo(ByVal arrRet)
	frm1.txtECNNo.Value			= arrRet(0)		
	frm1.txtECNNoDesc.Value		= arrRet(1)
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement	
End Function
'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
		
	frm1.txtChildItemCd.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetChildItemCd()  --------------------------------------------------
'	Name : SetChildItemCd()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetChildItemCd(byval arrRet)
	frm1.txtChildItemCd.Value    = arrRet(0)		
	frm1.txtChildItemNm.Value    = arrRet(1)
	
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Function

'=======================================================================================================
'   Event Name : txtChgFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtChgFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtChgFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtChgToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtChgToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtChgToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtChgFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtChgFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtChgToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtChgToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029  
	
	Call AppendNumberPlace("6", "4", "0")
	Call AppendNumberPlace("7", "2", "2")   
	Call AppendNumberPlace("8", "11", "6")	
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,FALSE,,ggStrMinPart,ggStrMaxPart)
	
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetToolbar("11000000000011")									'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row )
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("0000111111")

    If frm1.vspdData.MaxRows <= 0 Or Col < 1 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
   		If Col = C_InsertDt Then
   		   Col = C_InsertDtHD
   		End If

        ggoSpread.Source = frm1.vspdData
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
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

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
    
		If lgStrPrevKeyIndex <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False															'��: Processing is NG

    Err.Clear																	'��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
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
    End If     																		'��: Query db data

    FncQuery = True																'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
   Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)							'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)	                   '��:ȭ�� ����, Tab ���� 
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

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtBomNo=" & Trim(.txtBomNo.value) 
		strVal = strVal & "&txtChgFromDt=" & Trim(.txtChgFromDt.Text)		
		strVal = strVal & "&txtChgToDt=" & Trim(.txtChgToDt.Text)		
		strVal = strVal & "&txtECNNo=" & Trim(.txtECNNo.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtChildItemCd=" & Trim(.txtChildItemCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '��: Next key tag
    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	Call SetToolbar("11000000000111")									'��: ��ư ���� ���� 
    lgBlnFlgChgValue = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>BOM�̷���ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenChangeRef()">�����̷»�</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=90%>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=28 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>BOM Type</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBomNo" SIZE=5 MAXLENGTH=3 tag="12XXXU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBomNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBomNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtChgFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="11"> </OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtChgToDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="11"> </OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>���躯���ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNo" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="���躯���ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEcnCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNNo()">&nbsp;<INPUT TYPE=TEXT NAME="txtECNNoDesc" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>��ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChildItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="��ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChildItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenChildItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemNm" SIZE=20 tag="14"></TD>
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
								<TD HEIGHT="100%" COLSPAN = 4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hShiftCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
