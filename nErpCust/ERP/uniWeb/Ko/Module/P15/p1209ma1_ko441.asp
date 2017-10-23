<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Standard Routing
'*  3. Program ID           : P1204ma1
'*  4. Program Name         : Standard Routing Entry
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2002/12/03
'*  9. Modifier (First)     : Im Hyun Soo
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1209mb1_ko441.asp"											'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID= "p1209mb2_ko441.asp"											'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID = "p1209mb3_ko441.asp"											'��: �����Ͻ� ���� ASP�� 

Dim C_OprNo
Dim C_WcCd
Dim C_WcPopup
Dim C_WcNm
Dim C_JobCd
Dim C_JobNm
Dim C_InsideFlg
Dim C_InsideFlgDesc
Dim C_RoutOrder
Dim C_RoutOrderDesc

Dim C_BpCd
Dim C_BpPopup
Dim C_BpNm
Dim C_CurCd
Dim C_CurPopup
Dim C_SubconPrc
Dim C_TaxType
Dim C_TaxPopup
Dim C_MilestoneFlg

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
Dim lgChgValidToDtFlg
          
Dim BaseDate, StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_OprNo			= 1
	C_WcCd			= 2
	C_WcPopup		= 3
	C_WcNm			= 4
	C_JobCd			= 5
	C_JobNm			= 6
	C_InsideFlg		= 7
	C_InsideFlgDesc		= 8

	C_BpCd			= 9
	C_BpPopup		= 10
	C_BpNm			= 11
	C_CurCd			= 12
	C_CurPopup		= 13
	C_SubconPrc		= 14
	C_TaxType		= 15
	C_TaxPopup		= 16
	C_MilestoneFlg		= 17

	C_RoutOrder		= 18
	C_RoutOrderDesc		= 19
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    
    'lgChgValidToDtFlg = False
    
    lgIntGrpCount = 100                         'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                                       '��: initializes sort direction
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtValidFromDt.text = StartDate
	frm1.txtValidToDt.text =  UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	frm1.rdoMajorRouting1.checked = True

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        	frm1.txtPlantCd.value = lgPLCd
	End If
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
		ggoSpread.Spreadinit "V20091126",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_RoutOrderDesc + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0

		.Col = .MaxCols																'��: ������Ʈ�� ��� Hidden Column
		.ColHidden = True
    
		.Col = C_InsideFlg
		.ColHidden = True
    
		.Col = C_RoutOrder
		.ColHidden = True
		   
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit	C_OprNo, 		"����", 	 6,,,3,2
		ggoSpread.SSSetEdit	C_WcCd, 		"�۾���",	 8,,,7,2
		ggoSpread.SSSetButton	C_WcPopup
		ggoSpread.SSSetEdit	C_WcNm, 		"�۾����", 	20
		ggoSpread.SSSetCombo	C_JobCd, 		"�����۾��ڵ�",	12
		ggoSpread.SSSetCombo	C_JobNm, 		"�����۾���",	20
		ggoSpread.SSSetEdit	C_InsideFlg, 		"Ÿ��", 	 8
		ggoSpread.SSSetEdit	C_InsideFlgDesc, 	"Ÿ��", 	 8
		ggoSpread.SSSetCombo	C_RoutOrder, 	 	"�����ܰ�", 	10      
		ggoSpread.SSSetCombo	C_RoutOrderDesc, 	"�����ܰ�", 	10

		ggoSpread.SSSetEdit	C_BpCd, 		"����ó", 	10,,,18,2
		ggoSpread.SSSetButton 	C_BpPopup
		ggoSpread.SSSetEdit	C_BpNm, 		"����ó��", 	20
		ggoSpread.SSSetEdit	C_CurCd, 		"��ȭ", 	 6,,,3,2
		ggoSpread.SSSetButton 	C_CurPopup
		ggoSpread.SSSetFloat	C_SubconPrc,		"�������ִܰ�", 15, "C", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit	C_TaxType, 		"VAT����",	 8,,,5,2
		ggoSpread.SSSetButton	C_TaxPopup
		ggoSpread.SSSetCombo	C_MilestoneFlg, 	"Milestone", 	10

    
		Call ggoSpread.MakePairsColumn(C_WcCd, C_WcPopup)
		Call ggoSpread.MakePairsColumn(C_BpCd, C_BpPopup)
		Call ggoSpread.MakePairsColumn(C_CurCd, C_CurPopup)
		Call ggoSpread.MakePairsColumn(C_TaxType, C_TaxPopup)

		Call ggoSpread.SSSetColHidden(C_InsideFlg, C_InsideFlg, True)
		Call ggoSpread.SSSetColHidden(C_RoutOrder, C_RoutOrder, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SSSetSplit2(3)										'frozen ����߰� 
    
		.ReDraw = True
	
		Call SetSpreadLock 
    
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock	C_OprNo, -1,C_OprNo
    ggoSpread.SpreadLock	C_WcNm, -1,C_WcNm
    ggoSpread.spreadLock	C_InsideFlg, -1,C_InsideFlg
    ggoSpread.spreadLock	C_RoutOrder, -1,C_RoutOrder
    ggoSpread.spreadLock	C_InsideFlgDesc, -1,C_InsideFlgDesc    
    ggoSpread.spreadLock	C_RoutOrderDesc, -1,C_RoutOrderDesc

    ggoSpread.spreadLock	C_BpNm, -1, C_BpNm
    
    ggoSpread.SSSetRequired	C_WcCd, -1
    ggoSpread.SSSetProtected .vspdData.MaxCols, -1
    
    .vspdData.ReDraw = True

    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal InOutType)
    With frm1
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired		C_OprNo,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_WcCd,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_WcNm,		pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected	C_JobNm,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_InsideFlg,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_RoutOrder,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_InsideFlgDesc, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_RoutOrderDesc, pvStartRow, pvEndRow

		If InOutType = "N" Then

		   ggoSpread.SpreadUnLock	C_BpCd,		pvStartRow, C_BpPopup, pvEndRow
		   ggoSpread.SpreadUnLock	C_CurCd,	pvStartRow, C_CurPopup, pvEndRow
		   ggoSpread.SpreadUnLock	C_SubconPrc,	pvStartRow, C_SubconPrc, pvEndRow
		   ggoSpread.SpreadUnLock	C_TaxType,	pvStartRow, C_TaxPopup, pvEndRow

		   ggoSpread.SSSetRequired	C_BpCd,		pvStartRow, pvEndRow
		   ggoSpread.SSSetRequired	C_CurCd,	pvStartRow, pvEndRow
		   ggoSpread.SSSetRequired	C_SubconPrc,	pvStartRow, pvEndRow
		   ggoSpread.SSSetRequired	C_TaxType,	pvStartRow, pvEndRow
		   ggoSpread.SSSetProtected 	C_MilestoneFlg,	pvStartRow, pvEndRow
		Else
		   ggoSpread.SSSetProtected 	C_BpCd, 	pvStartRow, pvEndRow
		   ggoSpread.SSSetProtected 	C_BpPopup, 	pvStartRow, pvEndRow    
		   ggoSpread.SSSetProtected 	C_BpNm, 	pvStartRow, pvEndRow    
		   ggoSpread.SSSetProtected 	C_CurCd, 	pvStartRow, pvEndRow    
		   ggoSpread.SSSetProtected 	C_CurPopup,	pvStartRow, pvEndRow    
		   ggoSpread.SSSetProtected 	C_SubconPrc,	pvStartRow, pvEndRow
		   ggoSpread.SSSetProtected 	C_TaxType,	pvStartRow, pvEndRow
		   ggoSpread.SSSetProtected 	C_TaxPopup,	pvStartRow, pvEndRow    
		   ggoSpread.SSSetProtected 	C_RoutOrder, 	pvStartRow, pvEndRow
		   ggoSpread.SSSetRequired	C_MilestoneFlg,	pvStartRow, pvEndRow
		End If

		.vspdData.ReDraw = True
    End With
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
			C_OprNo			= iCurColumnPos(1)
			C_WcCd			= iCurColumnPos(2)
			C_WcPopup		= iCurColumnPos(3)
			C_WcNm			= iCurColumnPos(4)
			C_JobCd			= iCurColumnPos(5)
			C_JobNm			= iCurColumnPos(6)
			C_InsideFlg		= iCurColumnPos(7)
			C_InsideFlgDesc		= iCurColumnPos(8)

			C_BpCd			= iCurColumnPos(9)
			C_BpPopup		= iCurColumnPos(10)
			C_BpNm			= iCurColumnPos(11)
			C_CurCd			= iCurColumnPos(12)
			C_CurPopup		= iCurColumnPos(13)
			C_SubconPrc		= iCurColumnPos(14)
			C_TaxType		= iCurColumnPos(15)
			C_TaxPopup		= iCurColumnPos(16)
			C_MilestoneFlg		= iCurColumnPos(17)

			C_RoutOrder		= iCurColumnPos(18)
			C_RoutOrderDesc		= iCurColumnPos(19)

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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim iColSep
    Dim strCboCd
	
    iColSep = Parent.gColSep
	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    ggoSpread.Source = frm1.vspdData
	lgF0 = "" & iColSep & lgF0
	lgF1 = "" & iColSep & lgF1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobNm
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1201", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData
	lgF0 = "" & iColSep & lgF0
	lgF1 = "" & iColSep & lgF1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_RoutOrder
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_RoutOrderDesc

    strCboCd = ""
    strCboCd = "Y" & vbTab & "N"
    
    ggoSpread.SetCombo strCboCd, C_MilestoneFlg
    
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.Col = C_JobCd
			intIndex = .value
			.col = C_JobNm
			.value = intindex
	
			.Row = intRow
			.Col = C_RoutOrder
			intIndex = .value
			.col = C_RoutOrderDesc
			.value = intindex
			
		Next	
	End With
End Sub


Function SetFieldProp(ByVal lRow, ByVal sType)
	ggoSpread.Source = frm1.vspdData
	
	If sType = "N" Then			'���� �����̸� 
		ggoSpread.SpreadUnLock		C_BpCd,		lRow, C_BpPopup, lRow
		ggoSpread.SpreadUnLock		C_CurCd,	lRow, C_CurPopup, lRow
		ggoSpread.SpreadUnLock		C_SubconPrc,	lRow, C_SubconPrc, lRow
		ggoSpread.SpreadUnLock		C_TaxType,	lRow, C_TaxPopup, lRow

		ggoSpread.SSSetRequired		C_BpCd,		lRow, lRow
		ggoSpread.SSSetRequired		C_CurCd,	lRow, lRow
		ggoSpread.SSSetRequired		C_SubconPrc,	lRow, lRow
		ggoSpread.SSSetRequired		C_TaxType,		lRow, lRow

	ElseIf sType = "Y" Then		'�系 �����̸� 
		ggoSpread.SpreadLock		C_BpCd,		lRow, C_BpPopup,  lRow
		ggoSpread.SpreadLock		C_CurCd,	lRow, C_CurPopup, lRow
		ggoSpread.SpreadLock		C_SubconPrc,	lRow, C_SubconPrc, lRow
		ggoSpread.SpreadLock		C_TaxType,	lRow, C_TaxPopup, lRow

		ggoSpread.SSSetProtected	C_BpCd,		lRow, lRow
		ggoSpread.SSSetProtected	C_CurCd,	lRow, lRow
		ggoSpread.SSSetProtected	C_SubconPrc,	lRow, lRow
		ggoSpread.SSSetProtected	C_TaxType,	lRow, lRow
	End If
	
End Function
'------------------------------------------  OpenWcPopup()  -------------------------------------------------
'	Name : OpenWcPopup()
'	Description : WcPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWcPopup(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then		
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " AND VALID_TO_DT >=  " & FilterVar(BaseDate , "''", "S") & "" 
	arrParam(5) = "�۾���"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "HH" & parent.gcolsep & "INSIDE_FLG"
    arrField(3) = "CASE WHEN INSIDE_FLG=" & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("�系", "''", "S") & " ELSE " & FilterVar("����", "''", "S") & " END"
    arrField(4) = "dbo.ufn_GetCodeName(" & FilterVar("P1013", "''", "S") & ", WC_MGR)"
    
    arrHeader(0) = "�۾���"		
    arrHeader(1) = "�۾����"		
    arrHeader(2) = "�۾��屸��"		
    arrHeader(3) = "�۾��屸��"		
    arrHeader(4) = "�۾�������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWc(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData, C_WcCd, frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement

End Function

'------------------------------------------  OpenRouting()  -------------------------------------------------
'	Name : OpenRouting()
'	Description : RoutingPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenRouting_Old(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtRoutNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "ǥ�ض���� �˾�"	
	arrParam(1) = "(SELECT DISTINCT ROUT_NO, PLANT_CD, DESCRIPTION FROM P_STANDARD_ROUTING) A"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "ǥ�ض����"			
	
    arrField(0) = "ROUT_NO"
    arrField(1) = "DESCRIPTION"	
       
    arrHeader(0) = "ǥ�ض����"		
    arrHeader(1) = "ǥ�ض���ø�"		
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRouting(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function


Function OpenConRouting(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtRoutNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "ǰ��׷� �˾�"	
	arrParam(1) = "(SELECT DISTINCT ITEM_GROUP_CD, PLANT_CD, DESCRIPTION FROM P_ITEM_GROUP_STD_ROUTING_KO441 ) A"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "ǰ��׷�"			
	
    	arrField(0) = "ITEM_GROUP_CD"
    	arrField(1) = "DESCRIPTION"	
       
    	arrHeader(0) = "ǰ��׷�"		
    	arrHeader(1) = "ǰ��׷��"		
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRouting(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function


Function OpenRouting(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	Dim strSelect

	If IsOpenPop = True Or UCase(frm1.txtRoutNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True

	strSelect = ""
	strSelect = strSelect & " ( "

	strSelect = strSelect & "   Select ITEM_GROUP_CD, ITEM_GROUP_NM "
	strSelect = strSelect & "   FROM ( " 

	strSelect = strSelect & "          Select ITEM_GROUP_CD2 as ITEM_GROUP_CD, Max(ITEM_GROUP_NM2) as ITEM_GROUP_NM "
	strSelect = strSelect & "          From   VIEW_B_ITEM_GROUP_TREE2_KO441 "
	strSelect = strSelect & "          Where  ITEM_GROUP_CD2 <> '' "
       'strSelect = strSelect & "          and    ITEM_ACCT in ('10','20') "
	strSelect = strSelect & "          Group by ITEM_GROUP_CD2 "
	strSelect = strSelect & "        ) a "
	strSelect = strSelect & "    Where  a.ITEM_GROUP_CD not in (Select b.ITEM_GROUP_CD "  
	strSelect = strSelect & "                                    From   P_ITEM_GROUP_STD_ROUTING_KO441 b "
	strSelect = strSelect & "                                    Where  b.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	strSelect = strSelect & "                                    Group by b.ITEM_GROUP_CD "
	strSelect = strSelect & "                                   ) "
	strSelect = strSelect & " ) a "

	arrParam(0) = "ǰ��׷� �˾�"	
	arrParam(1) = strSelect
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "ǰ��׷�"			
	
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"	
       
	arrHeader(0) = "ǰ��׷�"		
	arrHeader(1) = "ǰ��׷��"		
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRouting2(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutingNo.focus
	
End Function


'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : OpenPlantPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    arrField(2) = "CUR_CD"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    arrHeader(2) = "��ȭ�ڵ�"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenBizPartner()  -------------------------------------------------
'	Name : OpenBizparener()
'	Description : BpPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizPartner(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó�˾�"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "����ó"			
	
    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"
    arrField(2) = "ED15" & parent.gcolsep & "BP_TYPE"
    arrField(3) = "ED15" & parent.gcolsep & "CURRENCY"
    arrField(4) = "ED15" & parent.gcolsep & "VAT_TYPE"
        
    arrHeader(0) = "BP"
    arrHeader(1) = "BP��"
    arrHeader(2) = "Bp ����"
    arrHeader(3) = "��ȭ"
    arrHeader(4) = "VAT����"
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetBizPartner(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_BpCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  SetBizPartner()  --------------------------------------------------
'	Name : SetBizPartner()
'	Description : RoutingNo Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizPartner(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_BpCd, .ActiveRow, UCase(arrRet(0))) 
		Call .SetText(C_BpNm, .ActiveRow, UCase(arrRet(1))) 
		Call .SetText(C_CurCd, .ActiveRow, UCase(arrRet(3))) 
		Call .SetText(C_TaxType, .ActiveRow, UCase(arrRet(4))) 
		Call vspdData_Change(0, .Row)	' ������ �Ͼ�ٰ� �˷��� 
	End With
End Function

'------------------------------------------  OpenCostCtr()  ----------------------------------------------
'	Name : OpenCostCtr()
'	Description : Cost Center Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCostCtr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(Frm1.txtCostCd.className) = "PROTECTED" Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	IsOpenPop = True 

	arrParam(0) = "Cost Center �˾�"			' �˾� ��Ī 
	arrParam(1) = "B_COST_CENTER"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtCostCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "B_COST_CENTER.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				" AND B_COST_CENTER.COST_TYPE ='M'" & _
				" AND B_COST_CENTER.DI_FG ='D'"			' Where Condition
	arrParam(5) = "Cost Center"					' TextBox ��Ī 
	
    arrField(0) = "COST_CD"							' Field��(0)
    arrField(1) = "COST_NM"							' Field��(1)
    
    arrHeader(0) = "Cost Center"				' Header��(0)
    arrHeader(1) = "Cost Center ��"				' Header��(1)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCostCtr(arrRet)
	End If	
    
End Function

'------------------------------------------  SetCostCtr()  -----------------------------------------------
'	Name : SetCostCtr()
'	Description : Cost Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetCostCtr(byval arrRet)
	frm1.txtCostCd.value = arrRet(0)
	frm1.txtCostNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  OpenCurrency()  -------------------------------------------------
'	Name : OpenCurrency()
'	Description : BpPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenCurrency(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ȭ�˾�"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "��ȭ"			
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
    
    arrHeader(0) = "��ȭ"		
    arrHeader(1) = "��ȭ��"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCurrency(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_CurCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  SetCurrency()  --------------------------------------------------
'	Name : SetCurrency()
'	Description : RoutingNo Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrency(Byval arrRet)
	With frm1
		.vspdData.Col = C_CurCd
		.vspdData.Text = UCase(arrRet(0))
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' ������ �Ͼ�ٰ� �˷��� 
	
	End With
End Function
'------------------------------------------  OpenVat()  -------------------------------------------------
'	Name : OpenVat()
'	Description : VAT popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenVat()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	frm1.vspdData.Col = C_TaxType
	If IsOpenPop = True Or UCase(frm1.vspdData.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VAT����"						
	arrParam(1) = "B_MINOR, B_CONFIGURATION"						
	
	arrParam(2) = Trim(frm1.vspdData.text)	
		
	arrParam(4) = "B_MINOR.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " AND B_MINOR.MINOR_CD=B_CONFIGURATION.MINOR_CD "							
	arrParam(4) = arrParam(4) & "AND B_MINOR.MAJOR_CD=B_CONFIGURATION.MAJOR_CD AND B_CONFIGURATION.SEQ_NO=1"
	arrParam(5) = "VAT����"							
	
    arrField(0) = "B_MINOR.MINOR_CD"					
    arrField(1) = "B_MINOR.MINOR_NM"
    arrField(2) = "F5" & parent.gColSep & "B_CONFIGURATION.REFERENCE"	
    
    arrHeader(0) = "VAT����"						
    arrHeader(1) = "VAT������"						
    arrHeader(2) = "VAT��"
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetVat(arrRet)
	End If	
End Function

Function SetVat(byval arrRet)
	frm1.vspdData.Col = C_TaxType
	frm1.vspdData.Text = arrRet(0)		
	Call vspdData_Change(frm1.vspdData.Col, frm1.vspdData.Row)		' ������ �Ͼ�ٰ� �˷��� 

	lgBlnFlgChgValue = True
End Function
'------------------------------------------  SetWc()  --------------------------------------------------
'	Name : SetWc()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWc(Byval arrRet)
	Dim lRow

	With frm1
		lRow = .vspdData.ActiveRow 

		.vspdData.Col = C_WcCd
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_WcNm
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_InsideFlg
		.vspdData.Text = UCase(arrRet(2))
		
		If UCase(arrRet(2)) = "Y" then
			.vspdData.Col = C_InsideFlgDesc 
			.vspdData.Text = "�系"

			.vspdData.Col = C_BpCd
			.vspdData.Text = ""
			.vspdData.Col = C_BpNm
			.vspdData.Text = ""			
			.vspdData.Col = C_CurCd
			.vspdData.Text = ""	
			.vspdData.Col = C_SubconPrc
			.vspdData.Text = ""		
			.vspdData.Col = C_TaxType
			.vspdData.Text = ""

			.vspdData.ReDraw = False

			Call SetFieldProp(lRow,"Y")
			
			.vspdData.ReDraw = True
		Else
			.vspdData.Col = C_InsideFlgDesc 
			.vspdData.Text = "����"

			.vspdData.Col = C_BpCd
			.vspdData.Text = "9999999999"
			.vspdData.Col = C_BpNm
			.vspdData.Text = "��Ÿ"			
			.vspdData.Col = C_CurCd
			.vspdData.Text = "KRW"	
			.vspdData.Col = C_SubconPrc
			.vspdData.Text = "1"		
			.vspdData.Col = C_TaxType
			.vspdData.Text = "A"

			.vspdData.ReDraw = False
			
			Call SetFieldProp(lRow,"N")
			
			.vspdData.ReDraw = True
		End if			
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' ������ �Ͼ�ٰ� �˷��� 
	
	End With
End Function

'------------------------------------------  SetBizPartner()  --------------------------------------------------
'	Name : SetBizPartner()
'	Description : RoutingNo Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizPartner(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_BpCd, .ActiveRow, UCase(arrRet(0))) 
		Call .SetText(C_BpNm, .ActiveRow, UCase(arrRet(1))) 
		Call .SetText(C_CurCd, .ActiveRow, UCase(arrRet(3))) 
		Call .SetText(C_TaxType, .ActiveRow, UCase(arrRet(4))) 
		Call vspdData_Change(0, .Row)	' ������ �Ͼ�ٰ� �˷��� 
	End With
End Function

'------------------------------------------  SetCurrency()  --------------------------------------------------
'	Name : SetCurrency()
'	Description : RoutingNo Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrency(Byval arrRet)
	With frm1
		.vspdData.Col = C_CurCd
		.vspdData.Text = UCase(arrRet(0))
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' ������ �Ͼ�ٰ� �˷��� 
	
	End With
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)	
	frm1.txtPlantNm.value	 = arrRet(1) 	
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetRouting()
'	Description : Routing Popup���� Routing NO setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRouting(byval arrRet)
	frm1.txtRoutNo.Value      = arrRet(0)		
	frm1.txtRoutingNm.Value   = arrRet(1)		
End Function

Function SetRouting2(byval arrRet)
	frm1.txtRoutingNo.Value  = arrRet(0)		
	frm1.txtRoutingNm1.Value  = arrRet(1)		
End Function


Sub LookUpBp(ByVal pBpCd)	'2003-08-29

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strSelect, strFrom, strWhere
	Dim arrRows0, arrRows1, arrRows2, arrRows3

	If Trim(pBpCd) = "" Then Exit Sub

	'----------------------------------------------------------------------------------
	strSelect	= " BP_CD, BP_NM, CURRENCY, VAT_TYPE "
	strFrom		= " B_BIZ_PARTNER "
	strWhere	= " BP_CD =  " & FilterVar(pBpCd, "''", "S") & " "
	
	With frm1.vspdData
	
	If CommonQueryRs(strSelect, strFrom, strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("189629", "X","X","X")

		Call .SetText(C_BpCd,	.ActiveRow, "")
		Call .SetText(C_BpNm,	.ActiveRow, "")
		'Call .SetText(C_CurCd,	.ActiveRow, "")
		'Call .SetText(C_TaxType,.ActiveRow, "")
	Else
		arrRows0 = Split(lgF0, Chr(11))
		arrRows1 = Split(lgF1, Chr(11))
		arrRows2 = Split(lgF2, Chr(11))
		arrRows3 = Split(lgF3, Chr(11))
		
		Call .SetText(C_BpCd,	.ActiveRow, arrRows0(0))
		Call .SetText(C_BpNm,	.ActiveRow, arrRows1(0))
		Call .SetText(C_CurCd,	.ActiveRow, arrRows2(0))
		Call .SetText(C_TaxType,.ActiveRow, arrRows3(0))
	End If
	
	End With
	
End Sub


Sub ProtectMilestone(ByVal pvFlag)
	Dim iIntCnt
	Dim iStrFlag
	Dim iStrMilestoneFlg
	Dim iStrInspFlg

	ggoSpread.SpreadUnLock 	C_MilestoneFlg, 1, C_MilestoneFlg, frm1.vspdData.MaxRows
	ggoSpread.SSSetRequired C_MilestoneFlg, 1, frm1.vspdData.MaxRows
	
	For iIntCnt = frm1.vspdData.MaxRows To 1 Step -1
		Call frm1.vspdData.GetText(0, iIntCnt, iStrFlag)
			
		Select Case iStrFlag
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
				Call frm1.vspdData.SetText(C_MilestoneFlg, iIntCnt, "Y")
				ggoSpread.SSSetProtected C_MilestoneFlg, iIntCnt, iIntCnt
				Exit For

			Case ""
				Call frm1.vspdData.GetText(C_MilestoneFlg, iIntCnt, iStrMilestoneFlg)

				If iStrMilestoneFlg = "N" Then
					Call frm1.vspdData.SetText(0, iIntCnt, ggoSpread.UpdateFlag)
				End If
				Call frm1.vspdData.SetText(C_MilestoneFlg, iIntCnt, "Y")

				ggoSpread.SSSetProtected C_MilestoneFlg, iIntCnt, iIntCnt
				Exit For
		End Select
	Next
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field    
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    '----------  Coding part  -------------------------------------------------------------
	
    Call InitComboBox
    
    Call SetToolbar("11101101001011")										'��: ��ư ���� ���� 
    
    Call GetValue_ko441()
    Call SetDefaultVal
    
    Call InitVariables
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtRoutNo.focus 
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
Sub vspdData_Click(ByVal Col , ByVal Row )
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("1101110111")

    If frm1.vspdData.MaxRows <= 0 Or Col < 1 Then
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
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	Select Case Col 
		Case  C_CurCd
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_CurCd,C_SubconPrc, "C" ,"X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_CurCd,C_SubconPrc, "C" ,"I","X","X")  
		Case  C_SubconPrc
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_CurCd,C_SubconPrc, "C" ,"X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_CurCd,C_SubconPrc, "C" ,"I","X","X")
		Case C_BpCd
			Call LookUpBp(Trim(GetSpreadText(frm1.vspdData, C_BpCd, Row,"X","X")))
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_CurCd,C_SubconPrc, "C" ,"X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_CurCd,C_SubconPrc, "C" ,"I","X","X")
	End Select	

    '----------  Coding part  -------------------------------------------------------------
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim strTemp
    Dim intPos1

	'----------  Coding part  -------------------------------------------------------------   
    With frm1.vspdData 
	
	ggoSpread.Source = frm1.vspdData

	If Row > 0 And Col = C_WcPopUp Then
	   .Col = C_WcCd
	   .Row = Row
        
	   Call OpenWcPopup(.Text)
        
	   Call SetActiveCell(frm1.vspdData,C_WcCd,Row,"M","X","X")
	   Set gActiveElement = document.activeElement

	ElseIf Row >0 And Col = C_BpPopup Then

	   .Col = C_BpCd
	   .Row = Row
        
	   Call OpenBizPartner(.Text)
	   Call SetActiveCell(frm1.vspdData,C_BpCd,Row,"M","X","X")
	   Set gActiveElement = document.activeElement
		
	ElseIf Row >0 And Col = C_CurPopup Then
		.Col = C_CurCd
		.Row = Row
        
		Call OpenCurrency(.Text)
		Call SetActiveCell(frm1.vspdData,C_CurCd,Row,"M","X","X")
		Set gActiveElement = document.activeElement
		
	ElseIf Row >0 And Col = C_TaxPopup Then
		.Col = C_TaxType
		.Row = Row
        
		Call OpenVAT()
		Call SetActiveCell(frm1.vspdData,C_TaxType,Row,"M","X","X")
		Set gActiveElement = document.activeElement
     
	End If
    End With

End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    If Row >= NewRow Then
        Exit Sub
    End If

	'----------  Coding part  -------------------------------------------------------------   

End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
    
		Select Case Col
			Case  C_JobCd
				.Col = Col
				intIndex = .Value
				.Col = C_JobNm
				.Value = intIndex
			Case  C_JobNm
				.Col = Col
				intIndex = .Value
				.Col = C_JobCd
				.Value = intIndex
		End Select
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
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtValidFromDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
	End If 
End Sub

'================================================ txtBox_onChange() ===============================
'   Event Name : txtBox_onChange()
'   Event Desc : 
'==========================================================================================
Sub rdoMajorRouting1_onClick()
	If lgRdoOldVal = 1 Then Exit Sub
	
	lgRdoOldVal = 1
	lgBlnFlgChgValue = True	
End Sub

Sub rdoMajorRouting2_onClick()
	If lgRdoOldVal = 2 Then Exit Sub
	
	lgRdoOldVal = 2
	lgBlnFlgChgValue = True	
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtValidToDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
	End If 
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================

Sub txtValidToDt_Change() 
	lgBlnFlgChgValue = True 
End Sub  


Function txtRoutNo_OnChange()
    Dim IntRetCd

    If  frm1.txtRoutNo.value = "" Then
        frm1.txtRoutingNm.Value = ""
    Else

        IntRetCD =  CommonQueryRs(" DISTINCT DESCRIPTION "," P_ITEM_GROUP_STD_ROUTING_KO441 (nolock) ", _
						    " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & "" & " AND ITEM_GROUP_CD = " & FilterVar(frm1.txtRoutNo.Value, "''", "S") & "" , _
						      lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        If IntRetCD=False   Then
	    frm1.txtRoutingNm.Value = ""
        Else
            frm1.txtRoutingNm.Value = Trim(Replace(lgF0,Chr(11),""))
        End If

     End If
     
End Function

Function txtRoutingNo_OnChange()
    Dim IntRetCd

    If  frm1.txtRoutingNo.value  = "" Then
        frm1.txtRoutingNm1.Value = ""
    Else

	if Trim(frm1.txtRoutingNm1.Value) = "" Then	
	   IntRetCD =  CommonQueryRs(" item_group_nm "," b_item_group (nolock) ", _
							" item_group_cd = " & FilterVar(frm1.txtRoutingNo.Value, "''", "S") & "" , _
							lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	   If IntRetCD=False  Then
		frm1.txtRoutingNm1.Value = ""
	   Else
                frm1.txtRoutingNm1.Value = Trim(Replace(lgF0,Chr(11),""))
	   End If
	End If
     End If
     
End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
    
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
    Call SetDefaultVal															'��: Initializes local global variables
    Call InitVariables

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
    End If     
    
    FncQuery = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    frm1.txtRoutNo.value = ""
    
    Call ggoOper.ClearField(Document, "2")											'��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    Call SetDefaultVal
	Call InitVariables																'��: Initializes local global variables
	
	Call SetToolbar("11101101001011")										'��: ��ư ���� ���� 
    
    frm1.txtRoutingNo.focus 
    Set gActiveElement = document.activeElement 
     
    FncNew = True																	'��: Processing is OK

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                  '��: �Ʒ� �޼����� DBȭ �ؼ� �� �������� ��ü 
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '��: "Will you destory previous data"	
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
	If DbDelete = False Then   
		Exit Function           
    End If         
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False AND ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '��: No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    
    ggoSpread.Source = frm1.vspdData
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		If frm1.vspdData.MaxRows = 0 Then
			Call DisplayMsgBox("971012", "X", "����", "X")
			Exit Function
		End If	
	End If
    
    If Not chkField(Document, "2") Then
		Exit Function
	End If
	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
   
    If DbSave = False Then   
		Exit Function           
    End If          '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
		
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement 
	frm1.vspdData.EditMode = True
	frm1.vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow, "N")
    
    frm1.vspdData.Col = C_OprNo
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    
    frm1.vspdData.Text = ""
    
    frm1.vspdData.ReDraw = True
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
	Call InitData(1)
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim iIntReqRows
    Dim iIntCnt

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

	With frm1
		.vspdData.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData
		.vspdData.EditMode = True
		.vspdData.ReDraw = False
    
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + iIntReqRows - 1,C_CurCd,C_SubconPrc, "C" ,"I","X","X")

		ggoSpread.InsertRow , iIntReqRows
        
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1, "Y")

		For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + iIntReqRows - 1
			.vspdData.Row = iIntCnt
			.vspdData.Col = C_MilestoneFlg
			.vspdData.Text = "N"    
		Next

		Call ProtectMilestone(0)
    
		.vspdData.ReDraw = True
    
    End With

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
    ggoSpread.Source = frm1.vspdData
    
    lDelRows = ggoSpread.DeleteRow

    Call ProtectMilestone(0)

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                               '��: Protect system from crashing
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
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

    DbQuery = False
    
    LayerShowHide(1)
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgCurDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtRoutNo=" & Trim(.txtRoutNo.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgCurDt=" & UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)														'��: ��ȸ ������ ������� 
    Dim lRow
    Dim InsideFlg

    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = false
    
    Call SetToolbar("11111111001111")
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

    Call InitData(LngMaxRow)												'��: Job Name Setting

    frm1.vspdData.ReDraw = False

    For lRow = LngMaxRow To frm1.vspdData.MaxRows

	frm1.vspdData.Col = C_InsideFlg
	frm1.vspdData.Row = lRow

	If UCase(Trim(frm1.vspdData.Text)) = "N" Then	'���� �����̸� 
		ggoSpread.SpreadUnLock		C_BpCd,		lRow, C_BpPopup, lRow
		ggoSpread.SpreadUnLock		C_CurCd,	lRow, C_CurPopup, lRow
		ggoSpread.SpreadUnLock		C_SubconPrc,	lRow, C_SubconPrc, lRow
		ggoSpread.SpreadUnLock		C_TaxType,	lRow, C_TaxPopup, lRow

		ggoSpread.SSSetRequired		C_BpCd,		lRow, lRow
		ggoSpread.SSSetRequired		C_CurCd,	lRow, lRow
		ggoSpread.SSSetRequired		C_SubconPrc,	lRow, lRow
		ggoSpread.SSSetRequired		C_TaxType,	lRow, lRow

	Else	'�系 �����̸� 
		ggoSpread.SpreadLock		C_BpCd,		lRow, C_BpPopup,  lRow
		ggoSpread.SpreadLock		C_CurCd,	lRow, C_CurPopup, lRow
		ggoSpread.SpreadLock		C_SubconPrc,	lRow, C_SubconPrc, lRow
		ggoSpread.SpreadLock		C_TaxType,	lRow, C_TaxPopup, lRow

		ggoSpread.SSSetProtected	C_BpCd,		lRow, lRow
		ggoSpread.SSSetProtected	C_CurCd,	lRow, lRow
		ggoSpread.SSSetProtected	C_SubconPrc,	lRow, lRow
		ggoSpread.SSSetProtected	C_TaxType,	lRow, lRow
	End IF	

    Next

    Call ProtectMilestone(1)

    frm1.vspdData.redraw = True

    If lgIntFlgMode <> parent.OPMD_UMODE Then
	Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
	Set gActiveElement = document.activeElement
    End If	

    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	
End Function


'========================================================================================
' Function Name : DbQuery
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
    Dim strMilestoneFlg, strInsideFlg
    Dim strSubconPrc

    DbSave = False                                                          '��: Processing is NG
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function      
	     
    LayerShowHide(1)
		
	With frm1
		
		.txtMode.value = parent.UID_M0002
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		iColSep = Parent.gColSep
		ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
		iValCnt = 0 : iDelCnt = 0
		
		'-----------------------
		'Data manipulate area
		'-----------------------

		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
 
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag							'��: �ű� 
		        
				strVal = ""
					
				strVal = strVal & "C" & iColSep & lRow & iColSep			'��: C=Create, Sheet�� 2�� �̹Ƿ� ����				                
		            
		               .vspdData.Col = C_OprNo			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep			'2
			            
			        .vspdData.Col = C_WCCd			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep			'3

			        .vspdData.Col = C_JobCd	
			        strVal = strVal & Trim(.vspdData.Text) & iColSep			'4

			        .vspdData.Col = C_InsideFlg
			        strVal = strVal & Trim(.vspdData.Text) & iColSep			'5
			        strInsideFlg = Trim(.vspdData.Text)
			       
				.vspdData.Col = C_RoutOrder
			        strVal = strVal & Trim(.vspdData.Text) & iColSep			'6
			        
			        strVal = strVal & UNIConvDate(.txtValidFromDt.Text) & iColSep		'7
			        strVal = strVal & UNIConvDate(.txtValidToDt.Text) & iColSep		'8
			        strVal = strVal & .txtRoutingNm1.value & iColSep			'9


				.vspdData.Col = C_MilestoneFlg						'10
				strVal = strVal & Trim(.vspdData.Text) & iColSep
				strMilestoneFlg = Trim(.vspdData.Text)

				If strInsideFlg = "N" Then
				   .vspdData.Col = C_BpCd
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'11

				   .vspdData.Col = C_SubconPrc
				   strSubconPrc = Trim(.vspdData.Text)
				   If strInsideFlg = "N" And UNIConvNum(strSubconPrc,0) = 0 Then
					Call DisplayMsgBox("970022", "X" , "�������ִܰ�", "0")
					Call SheetFocus(lRow, C_SubconPrc)
					Exit Function
				   End If
				   strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep	'12

				   .vspdData.Col = C_CurCd			
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'13

				   .vspdData.Col = C_TaxType				
				   strVal = strVal & UCase(Trim(.vspdData.Text)) & parent.gRowSep	'14

				Else
				   strVal = strVal & "" & iColSep					'11
				   strVal = strVal & "0" & iColSep					'12
				   strVal = strVal & "" & iColSep					'13
				   strVal = strVal & "" & parent.gRowSep				'14
				End If


			        ReDim Preserve TmpBufferVal(iValCnt)
			        TmpBufferVal(iValCnt) = strVal
			        iValCnt = iValCnt + 1

		        Case ggoSpread.UpdateFlag											'��: �ű� 
		        
				strVal = ""
					
				strVal = strVal & "U" & iColSep & lRow & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ����				                
		            
		            	.vspdData.Col = C_OprNo			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep		'2
			            
			        .vspdData.Col = C_WCCd			
			        strVal = strVal & Trim(.vspdData.Text) & iColSep		'3

			        .vspdData.Col = C_JobCd	
			        strVal = strVal & Trim(.vspdData.Text) & iColSep		'4

			        .vspdData.Col = C_InsideFlg
			        strVal = strVal & Trim(.vspdData.Text) & iColSep		'5
			        strInsideFlg = Trim(.vspdData.Text)

			        .vspdData.Col = C_RoutOrder
			        strVal = strVal & Trim(.vspdData.Text) & iColSep		'6
			        
			        strVal = strVal & UNIConvDate(.txtValidFromDt.Text) & iColSep	'7
			        strVal = strVal & UNIConvDate(.txtValidToDt.Text) & iColSep	'8
			        strVal = strVal & .txtRoutingNm1.value & iColSep		'9

				.vspdData.Col = C_MilestoneFlg						'10
				strVal = strVal & Trim(.vspdData.Text) & iColSep
				strMilestoneFlg = Trim(.vspdData.Text)

				If strInsideFlg = "N" Then
				   .vspdData.Col = C_BpCd
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'11

				   .vspdData.Col = C_SubconPrc
				   strSubconPrc = Trim(.vspdData.Text)
				   If strInsideFlg = "N" And UNIConvNum(strSubconPrc,0) = 0 Then
					Call DisplayMsgBox("970022", "X" , "�������ִܰ�", "0")
					Call SheetFocus(lRow, C_SubconPrc)
					Exit Function
				   End If
				   strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep	'12

				   .vspdData.Col = C_CurCd			
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'13

				   .vspdData.Col = C_TaxType				
				   strVal = strVal & UCase(Trim(.vspdData.Text)) & parent.gRowSep	'14

				Else
				   strVal = strVal & "" & iColSep					'11
				   strVal = strVal & "0" & iColSep					'12
				   strVal = strVal & "" & iColSep					'13
				   strVal = strVal & "" & parent.gRowSep				'14
				End If
			        
			        ReDim Preserve TmpBufferVal(iValCnt)
			        TmpBufferVal(iValCnt) = strVal
			        iValCnt = iValCnt + 1
		            
		        Case ggoSpread.DeleteFlag											'��: ���� 
					
				strDel = ""
					
				strDel = strDel & "D" & iColSep & lRow & iColSep		'��: D=Delete
					
				.vspdData.Col = C_OprNo						'2
				strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep									
		            
				ReDim Preserve TmpBufferDel(iDelCnt)
				TmpBufferDel(iDelCnt) = strDel
				iDelCnt = iDelCnt + 1
		            
			Case Else
				If lgBlnFlgChgValue = True Then
						
				   strVal = ""
						
				   strVal = strVal & "U" & iColSep & lRow & iColSep			'��: U=Update		
			
				   .vspdData.Col = C_OprNo
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'2
			            
				   .vspdData.Col = C_WCCd			
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'3

				   .vspdData.Col = C_JobCd	
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'4

				   .vspdData.Col = C_InsideFlg
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'5
			           strInsideFlg = Trim(.vspdData.Text)

				   .vspdData.Col = C_RoutOrder
				   strVal = strVal & Trim(.vspdData.Text) & iColSep			'6
			        
				   strVal = strVal & UNIConvDate(.txtValidFromDt.Text) & iColSep	'7
				   strVal = strVal & UNIConvDate(.txtValidToDt.Text) & iColSep		'8
				   strVal = strVal & .txtRoutingNm1.value & iColSep			'9


				   .vspdData.Col = C_MilestoneFlg					'10
				   strVal = strVal & Trim(.vspdData.Text) & iColSep
				   strMilestoneFlg = Trim(.vspdData.Text)

				   If strInsideFlg = "N" Then
				   	.vspdData.Col = C_BpCd
				   	strVal = strVal & Trim(.vspdData.Text) & iColSep		'11

				   	.vspdData.Col = C_SubconPrc
				   	strSubconPrc = Trim(.vspdData.Text)
				   	If strInsideFlg = "N" And UNIConvNum(strSubconPrc,0) = 0 Then
					   Call DisplayMsgBox("970022", "X" , "�������ִܰ�", "0")
					   Call SheetFocus(lRow, C_SubconPrc)
					   Exit Function
				   	End If
				   	strVal = strVal & UNIConvNum(Trim(.vspdData.Text), 0) & iColSep	'12

				   	.vspdData.Col = C_CurCd			
				   	strVal = strVal & Trim(.vspdData.Text) & iColSep		'13

				   	.vspdData.Col = C_TaxType				
				  	strVal = strVal & UCase(Trim(.vspdData.Text)) & parent.gRowSep	'14

				   Else
				   	strVal = strVal & "" & iColSep					'11
				   	strVal = strVal & "0" & iColSep					'12
				  	strVal = strVal & "" & iColSep					'13
					strVal = strVal & "" & parent.gRowSep				'14
				   End If


				   ReDim Preserve TmpBufferVal(iValCnt)
				   TmpBufferVal(iValCnt) = strVal
				   iValCnt = iValCnt + 1
			        
				End If
		    End Select
		            
		Next
		
		iTotalStrDel = Join(TmpBufferDel, "")
		iTotalStrVal = Join(TmpBufferVal, "")
 		
		.txtSpread.value = iTotalStrDel & iTotalStrVal
		
		.txtMaxRows.value = .vspdData.MaxRows
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
	frm1.txtRoutNo.value = frm1.txtRoutingNo.value 
	
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete()
	Dim strVal
	
	DbDelete = False														'��: Processing is NG
	
	LayerShowHide(1)
		
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'��: ���� ���� ����Ÿ 
    strVal = strVal & "&txtRoutingNo=" & Trim(frm1.txtRoutingNo.value)				'��: ���� ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    DbDelete = True                                                         '��: Processing is NG 
End Function
Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
End Function

Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ��׷�ǥ�ض���õ��</font></td>
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
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12NXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant frm1.txtPlantCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>ǰ��׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=15 MAXLENGTH=7 tag="12XXXU" ALT = "ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConRouting frm1.txtRoutNo.value">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm" SIZE=30 MAXLENGTH=50 tag="14"></TD>
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
								<TD CLASS="TD5" NOWRAP>ǰ��׷�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutingNo" SIZE=20 MAXLENGTH=7 tag="23XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRouting frm1.txtRoutingNo.value">&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm1" SIZE=30 MAXLENGTH=50 tag="22" ALT="ǰ��׷��"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>��ȿ�Ⱓ</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/p1204ma1_I545670807_txtValidFromDt.js'></script>&nbsp;~&nbsp;
									<script language =javascript src='./js/p1204ma1_I342035514_txtValidToDt.js'></script>
								</TD>
								<TD CLASS=TD5 NOWRAP>�۾����� C/C</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" SIZE=15 MAXLENGTH=10 tag="23XXXU" ALT="�۾����� C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCtr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCtr()">&nbsp;<INPUT NAME="txtCostNm" MAXLENGTH="20" SIZE=30 ALT ="�ڽ�Ʈ��Ÿ��" tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ֶ����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoMajorRouting" ID="rdoMajorRouting1" Value="Y" CLASS="RADIO" tag="23X" CHECKED><LABEL FOR="rdoMajorRouting1">��</LABEL>
													 <INPUT TYPE="RADIO" NAME="rdoMajorRouting" ID="rdoMajorRouting2" Value="N" CLASS="RADIO" tag="23X"><LABEL op="rdoMajorRouting2">�ƴϿ�</LABEL></TD>
								<TD CLASS=TD5 NOWRAP>����� ����</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID="CLSID:DD55D13D-EBF7-11D0-8810-0000C0E5948C" name=txtALTRTVALUE CLASS=FPDS65 title=FPDOUBLESINGLE SIZE="3" MAXLENGTH="3" ALT="����" tag="23X6Z" id=OBJECT1> </OBJECT>');</SCRIPT>
								</TD>							
							</TR>	
							<TR>
								<TD HEIGHT="100%" COLSPAN = 4>
								<script language =javascript src='./js/p1204ma1_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
