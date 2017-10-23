<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Component Allocation Entry
'*  3. Program ID           : p1205ma1.asp
'*  4. Program Name         : Entry Bill Of Resource
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2000/03/14
'*  8. Modified date(Last)  : 2002/12/18
'*  9. Modifier (First)     : Mr  Kim GyoungDon
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p1205mb1.asp"								'☆: 비지니스 로직 ASP명 

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p1205mb2.asp"								'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "p1205mb3.asp"								'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_LOOKUP_ID = "p1205mb4.asp"

' Grid 1(vspdData1) - Operation 
Dim C_OprNo
Dim C_WCCd
Dim C_JobCd
Dim C_JobNm
Dim C_InsideFlg
Dim C_MfgLt
Dim C_QueueTime
Dim C_SetupTime
Dim C_WaitTime
Dim C_FixRunTime
Dim C_RunTime
Dim C_ItemQtyForRunTime
Dim C_UnitOfItemQtyForRunTime
Dim C_MoveTime
Dim C_OverlapOpr
Dim C_OverlapLt
Dim C_BpCd
Dim C_CurCd
Dim C_UnitPriceOfOprSubcon
Dim C_TaxType
Dim C_MilestoneFlg
Dim C_RoutOrder
Dim C_ValidFromDt
Dim C_ValidToDt

' Grid 2(vspdData2) - Operation 
Dim C_Rank
Dim C_ResourceCd
Dim C_ResourcePopup
Dim C_ResourceNm
Dim C_ResourceType
Dim C_Efficiency

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgStrPrevKey2
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow1
Dim lgOldRow2
Dim lgSortKey1
Dim lgSortKey2

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvGridId)
	If pvGridId = "*" Or pvGridId = "A" Then
		' Grid 1(vspdData1) - Operation 
		C_OprNo					= 1
		C_WCCd					= 2
		C_JobCd					= 3
		C_JobNm					= 4
		C_InsideFlg				= 5
		C_MfgLt					= 6
		C_QueueTime				= 7
		C_SetupTime				= 8
		C_WaitTime				= 9
		C_FixRunTime			= 10
		C_RunTime				= 11
		C_ItemQtyForRunTime		= 12
		C_UnitOfItemQtyForRunTime = 13
		C_MoveTime				= 14
		C_OverlapOpr			= 15
		C_OverlapLt				= 16
		C_BpCd					= 17
		C_CurCd					= 18
		C_UnitPriceOfOprSubcon	= 19
		C_TaxType				= 20
		C_MilestoneFlg			= 21
		C_RoutOrder				= 22
		C_ValidFromDt			= 23
		C_ValidToDt				= 24
	End If

	If pvGridId = "*" Or pvGridId = "B" Then
		' Grid 2(vspdData2) - Operation 
		C_Rank				= 1
		C_ResourceCd		= 2
		C_ResourcePopup		= 3
		C_ResourceNm		= 4
		C_ResourceType		= 5
		C_Efficiency		= 6
	End If
End Sub
         
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey2 = ""
    lgStrPrevKey = ""							'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgOldRow1 = 0
    lgOldRow2 = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value	= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value	= ReadCookie("txtPlantNm")
		frm1.txtItemCd.Value	= ReadCookie("txtItemCd")
		frm1.txtItemNm.value	= ReadCookie("txtItemNm")
		frm1.txtRoutNo.Value	= ReadCookie("txtRoutingNo")
		frm1.txtRoutNm.value	= ReadCookie("txtRoutingNm")
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtRoutingNo", ""
	WriteCookie "txtRoutingNm", ""
	
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
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 

			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

			.ReDraw = False

			.MaxCols = C_ValidToDt + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
				
			.Col = C_TaxType		
			.ColHidden = True
			        
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit	C_OprNo,				"공정", 7,,,3,2
			ggoSpread.SSSetEdit	C_WCCd,					"작업장", 10,,,7,2
			ggoSpread.SSSetCombo	C_JobCd,			"공정작업", 10
			ggoSpread.SSSetCombo	C_JobNm,			"공정작업명", 10
			ggoSpread.SSSetEdit	C_InsideFlg,			"공정타입", 10
			ggoSpread.SSSetEdit	C_MfgLt,				"제조L/T", 7,1,,3
			ggoSpread.SSSetTime	C_QueueTime,			"Queue시간", 10,2 ,1 ,1
			ggoSpread.SSSetTime	C_SetupTime,			"설치시간", 10,2 ,1 ,1
			ggoSpread.SSSetTime	C_WaitTime,				"대기시간", 10,2 ,1 ,1
			ggoSpread.SSSetTime	C_FixRunTime,				"고정가동시간", 10,2 ,1 ,1
			ggoSpread.SSSetTime	C_RunTime,				"변동가동시간", 10,2 ,1 ,1
			ggoSpread.SSSetFloat	C_ItemQtyForRunTime,"기준수량", 15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_UnitOfItemQtyForRunTime, "기준단위", 10,,,3,2
			ggoSpread.SSSetTime	C_MoveTime,				"이동시간", 10,2 ,1 ,1
			ggoSpread.SSSetEdit	C_OverLapOpr,			"Overlap 공정", 7,,,3,2
			ggoSpread.SSSetEdit	C_OverLapLt,			"Overlap L/T",8,1
			ggoSpread.SSSetEdit	C_BpCd,					"외주처", 10,,,18,2
			ggoSpread.SSSetEdit	C_CurCd,				"통화", 6,,,3,2
			'ggoSpread.SSSetFloat	C_UnitPriceOfOprSubcon,	"공정외주단가",15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_UnitPriceOfOprSubcon,	"공정외주단가",15,"C"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit	C_TaxType,				"VAT유형", 6,,,5,2
			ggoSpread.SSSetEdit	C_MilestoneFlg,			"Milestone", 10
			ggoSpread.SSSetEdit	C_RoutOrder,			"공정단계", 10
			ggoSpread.SSSetDate 	C_ValidFromDt,		"시작일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ValidToDt,		"종료일", 11, 2, parent.gDateFormat
				
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

			ggoSpread.SSSetSplit2(2)										'frozen 기능추가 
				
			.ReDraw = true
    
		End With
	End If
	
	If pvGridId = "*" Or pvGridId = "B" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
	
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021126",, parent.gAllowDragDropSpread

			.ReDraw = False

			.MaxCols = C_Efficiency + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
	
			Call GetSpreadColumnPos("B")
	
			ggoSpread.SSSetFloat	C_Rank,			"순서", 7, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ResourceCd,	"자원", 20,,,10,2
			ggoSpread.SSSetButton 	C_ResourcePopup
			ggoSpread.SSSetEdit		C_ResourceNm,	"자원명", 30 
			ggoSpread.SSSetEdit		C_ResourceType,	"자원구분",20
			ggoSpread.SSSetFloat	C_Efficiency,	"효율", 10,	"7", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

			Call ggoSpread.MakePairsColumn(C_ResourceCd, C_ResourcePopup)
			Call ggoSpread.SSSetColHidden(C_Rank, C_Rank, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			ggoSpread.SSSetSplit2(3)										'frozen 기능추가 

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
	    '--------------------------------
	    'Grid 1
	    '--------------------------------
	    ggoSpread.Source = frm1.vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If

	If pvGridId = "*" Or pvGridId = "B" Then
		    
	    '--------------------------------
	    'Grid 2
	    '--------------------------------
	    ggoSpread.Source = frm1.vspdData2
	    .vspdData2.ReDraw = False
	
		ggoSpread.SpreadLock C_Rank,		-1,	C_Rank
		ggoSpread.SpreadLock C_ResourceCd,	-1,	C_ResourceCd
		ggoSpread.SpreadLock C_ResourcePopup,-1,C_ResourcePopup
		ggoSpread.SpreadLock C_ResourceNm,	-1,	C_ResourceNm
		ggoSpread.SpreadLock C_ResourceType,-1,	C_ResourceType
		ggoSpread.SpreadUnLock C_Efficiency,	-1
		
		ggoSpread.SSSetProtected	.vspdData2.MaxCols, -1
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
		ggoSpread.SSSetRequired		C_ResourceCd,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_Rank,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ResourceType, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ResourceNm, 	pvStartRow, pvEndRow
			   
		.Col = 1
		.Row = .ActiveRow
		.Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
		.EditMode = True
	   
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
       Case "A"
			' Grid 1(vspdData1) - Operation 
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_OprNo					= iCurColumnPos(1)
			C_WCCd					= iCurColumnPos(2)
			C_JobCd					= iCurColumnPos(3)
			C_JobNm					= iCurColumnPos(4)
			C_InsideFlg				= iCurColumnPos(5)
			C_MfgLt					= iCurColumnPos(6)
			C_QueueTime				= iCurColumnPos(7)
			C_SetupTime				= iCurColumnPos(8)
			C_WaitTime				= iCurColumnPos(9)
			C_FixRunTime			= iCurColumnPos(10)
			C_RunTime				= iCurColumnPos(11)
			C_ItemQtyForRunTime		= iCurColumnPos(12)
			C_UnitOfItemQtyForRunTime = iCurColumnPos(13)
			C_MoveTime				= iCurColumnPos(14)
			C_OverlapOpr			= iCurColumnPos(15)
			C_OverlapLt				= iCurColumnPos(16)
			C_BpCd					= iCurColumnPos(17)
			C_CurCd					= iCurColumnPos(18)
			C_UnitPriceOfOprSubcon	= iCurColumnPos(19)
			C_TaxType				= iCurColumnPos(20)
			C_MilestoneFlg			= iCurColumnPos(21)
			C_RoutOrder				= iCurColumnPos(22)
			C_ValidFromDt			= iCurColumnPos(23)
			C_ValidToDt				= iCurColumnPos(24)

       Case "B"
			' Grid 2(vspdData2) - Operation 
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_Rank			= iCurColumnPos(1)
			C_ResourceCd	= iCurColumnPos(2)
			C_ResourcePopup	= iCurColumnPos(3)
			C_ResourceNm	= iCurColumnPos(4)
			C_ResourceType	= iCurColumnPos(5)
			C_Efficiency	= iCurColumnPos(6)
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
    Call InitComboBox(gActiveSpdSheet.id)
	Call ggoSpread.ReOrderingSpreadData()
	If gActiveSpdSheet.id = "A" Then
		Call InitData(1)
	End If
End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitComboBox(ByVal pvGridId)

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	If pvGridId = "*" Or pvGridId = "A" Then
		Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		ggoSpread.Source = frm1.vspdData1
		lgF0 = "" & parent.gColSep & lgF0
		lgF1 = "" & parent.gColSep & lgF1
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobCd
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobNm
	End If
	
End Sub

'------------------------------------------  OpenConItemInfo()  -------------------------------------------------
'	Name : OpenConItemInfo()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X" , "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
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
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenConPlant()  -----------------------------------------
'	Name : OpenConPlant()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

Function OpenConRouting()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtRoutNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "품목", "X")
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "라우팅팝업"	
	arrParam(1) = "P_ROUTING_HEADER"				
	arrParam(2) = Trim(frm1.txtRoutNo.Value)
	arrParam(3) = ""
	arrParam(4) =  "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " And ITEM_CD = " & FilterVar(frm1.txtItemCd.value, "''", "S")
	arrParam(5) = "라우팅"			

    arrField(0) = "ROUT_NO"	
    arrField(1) = "DESCRIPTION"	
    arrField(2) = "BOM_NO"
    arrField(3) = "MAJOR_FLG"

    arrHeader(0) = "라우팅"		
    arrHeader(1) = "라우팅명"		
    arrHeader(2) = "BOM Type"
    arrHeader(3) = "주라우팅"
    
    	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRouting(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource(ByVal pVal)
	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "자원팝업"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(pVal)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "자원"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    arrField(2) = "dbo.ufn_GetCodeName(" & FilterVar("p1502", "''", "S") & ",RESOURCE_TYPE)"	
    
    arrHeader(0) = "자원"		
    arrHeader(1) = "자원명"
    arrHeader(2) = "자원구분"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If
	
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
	End With
End Function

'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)
End Function
'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	frm1.vspdData2.Col = C_ResourceCd
	frm1.vspdData2.Text = arrRet(0)
	frm1.vspdData2.Col = C_ResourceNm
	frm1.vspdData2.Text = arrRet(1)
	frm1.vspdData2.Col = C_ResourceType
	frm1.vspdData2.Text = arrRet(2)
	lgBlnFlgChgValue = True	
End Function
'------------------------------------------  SetRouting()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRouting(byval arrRet)
	frm1.txtRoutNo.Value    = arrRet(0)
	frm1.txtRoutNm.Value    = arrRet(1)
End Function

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData1
		For intRow = 1 To .MaxRows
			.Row = intRow
			.Col = C_JobCd
			intIndex = .value
			.col = C_JobNm
			.value = intindex
		Next	
	End With
End Sub

Function LookUpResource()
	Dim strVal
	
	LayerShowHide(1) 
		
	Err.Clear
	With frm1		
		strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)						
		strVal = strVal & "&txtResourceCd=" & Trim(.vspdData2.Text) 		
		strVal = strVal & "&lgLngCurRows=" & .vspdData2.ActiveRow 		
	End With

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet("*")                                                    '⊙: Setup the Spread sheet
    
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call InitComboBox("*")
    Call SetToolbar("11000000000011")
    
    If parent.gPlant <> "" And frm1.txtPlantCd.value = "" Then
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

'=======================================================================================================
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData1_onfocus()
End Sub

'=======================================================================================================
'   Event Name : vspdData2_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData2_onfocus()
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	Dim IntRetCD
	
    Call SetPopupMenuItemInf("0000110111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData1
	
	If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
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
 		
 		lgOldRow1 = Row

		frm1.vspdData2.MaxRows = 0
		
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_OprNo
			
		frm1.hOprNo.value = Trim(frm1.vspdData1.Text) 
			
		If DbDtlQuery = False Then		
			Call RestoreToolBar()
			Exit Sub
		End If
	Else
		If lgOldRow1 <> Row Then
			
			ggoSpread.Source = frm1.vspdData2

			If ggoSpread.SSCheckChange = True Then
				IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
				If IntRetCD = vbNo Then
					Exit Sub
				End If
			End If
			
			frm1.vspdData1.Row = Row
			frm1.vspdData1.Col = C_OprNo
			
			frm1.hOprNo.value = Trim(frm1.vspdData1.Text) 
	
			lgOldRow1 = Row
			
			frm1.vspdData2.MaxRows = 0
			
			LayerShowHide(1) 
				
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbDtlQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If	
			
		End If
	End If	
End Sub


'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2

	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1101110111")
	Else
		Call SetPopupMenuItemInf("0000110111")
	End If

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
 	End If
 	
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

'=======================================================================================================
'   Event Name : vspdData2_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	
	If Row <= 0 Or Col < 0 Then
		Exit Sub
	End If
	If NewRow <= 0 or NewCol <= 0 Then
		Exit Sub
	End If
	If lgOldRow2 <> Row Then
		lgLngCurRows = NewRow
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData2_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.UpdateRow Row

	With frm1.vspdData2
	
		.Row = Row
    
		Select Case Col
			Case  C_ResourceCd						
				.Col = Col
				If .Text <> "" Then
					Call LookUpResource
				Else
					.Col = C_ResourceNm
					.Text = ""
					.Col = C_ResourceType
					.Text = ""
				End If
		End Select
    
    End With
	
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

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

	With frm1.vspdData2 
	
    ggoSpread.Source = frm1.vspdData2
   
    If Row > 0 And Col = C_ResourcePopup Then
        .Col = C_ResourceCd
        .Row = Row
        
        Call OpenResource(.Text)
        Call SetActiveCell(frm1.vspdData2,C_ResourceCd,Row,"M","X","X")
		Set gActiveElement = document.activeElement
        
    End If
    
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
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1, NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then
		If lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			LayerShowHide(1) 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbDtlQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If									'⊙: 작업진행중 표시	
			
		End If     
    End if
    
End Sub


'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData2_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData2
	
		.Row = Row
    
		'Select Case Col
			'Case  C_LoadTypeNm
		'		.Col = Col
		'		intIndex = .Value
		'		.Col = C_LoadType
		'		.Value = intIndex
		'End Select
    
    End With

End Sub


Sub txtPlantCd_OnChange()
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If	
End Sub

Sub txtItemCd_OnChange()
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If	
End Sub

Sub txtRoutNo_OnChange()
	If frm1.txtRoutNo.value = "" Then
		frm1.txtRoutNm.value = ""
	End If	
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False															'⊙: Processing is NG
    
    Err.Clear																	'☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2 
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
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
	
	If frm1.txtRoutNo.value = "" Then
		frm1.txtRoutNm.value = ""
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     													'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
    	
    If ggoSpread.SSCheckChange = False Then 
       IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
       Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		LayerShowHide(0) 
		Exit Function           
    End If     				                                                  '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    
    If frm1.vspdData2.MaxRows < 1 Then Exit Function
    
	lgBlnFlgChgValue = TRUE
	
	frm1.vspdData2.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData2	
    ggoSpread.CopyRow
    Call SetSpreadColor(frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow)
    
	frm1.vspdData2.ReDraw = True
	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	If frm1.vspdData2.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData2	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt)
    Dim iIntReqRows, iIntCnt
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If

    With frm1.vspdData2
 
		.Focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = frm1.vspdData2
		.EditMode = True
    
		.ReDraw = False

		ggoSpread.InsertRow , iIntReqRows
		    
        Call SetSpreadColor(.ActiveRow, .ActiveRow + iIntReqRows - 1)
    
		lgLngCurRows = .ActiveRow
		For iIntCnt = .ActiveRow To .ActiveRow + iIntReqRows - 1
			
			.Row = iIntCnt
			.Col = C_Efficiency
			.Value = 100
			.Col = C_Rank
			.Text = 0
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
    Dim lTempRows 

    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData2.maxrows < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData2
    lDelRows = ggoSpread.DeleteRow
    
    lgBlnFlgChgValue = True

End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                                    '☜: Protect system from crashing
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
	ggoSpread.Source = frm1.vspdData2
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
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
   
    LayerShowHide(1) 
		
    Err.Clear                                                               '☜: Protect system from crashing
        
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
    Else
	
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtRoutNo=" & Trim(.txtRoutNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Call InitData(LngMaxRow)

	frm1.vspdData1.Col = C_OprNo
	frm1.vspdData1.Row = 1
	
	frm1.hOprNo.value = Trim(frm1.vspdData1.Text) 
	
	lgOldRow1 = 1
		
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisableToolBar(parent.TBC_QUERY)
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		  
		If DbDtlQuery = False Then
			Call RestoreToolBar()
			Exit Function
		End If				
		
	End If

	lgIntFlgMode = parent.OPMD_UMODE
	Call SetToolbar("1100111100111111")
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbDtlQuery() 
    Dim strVal
    
    DbDtlQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.value)
		strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.value)
		strVal = strVal & "&txtBomNo=" & Trim(.txtBomNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
    Else
				
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtRoutNo=" & Trim(.txtRoutNo.value)
		strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.value)
		strVal = strVal & "&txtBomNo=" & Trim(.txtBomNo.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
    End If

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel
	Dim ChkTimeVal, strQuantity, strHdnQty, strStatus
	Dim iColSep
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt
	
    DbSave = False                                                          
    
    LayerShowHide(1) 
		
    On Error Resume Next                                                   

	With frm1
		.txtMode.value = parent.UID_M0002
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    iValCnt = 0 : iDelCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData2.MaxRows
    
		ggoSpread.Source = .vspdData2 
    
        .vspdData2.Row = lRow
        .vspdData2.Col = 0
        
		strStatus = Trim(.vspdData2.Text)
		Select Case strStatus
		    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
				strVal = ""
				If .vspdData2.Text = ggoSpread.InsertFlag Then
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
				ElseIf .vspdData2.Text = ggoSpread.UpdateFlag Then
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep				'⊙: U=Update
				End If	
		        .vspdData2.Col = C_ResourceCd						'2
		        strVal = strVal & UCase(Trim(.vspdData2.Text)) & parent.gColSep
		        .vspdData2.Col = C_Rank								'3
		        strVal = strVal & UniConvNum(.vspdData2.Text,0) & parent.gColSep
		        .vspdData2.Col = C_Efficiency						'4
				strVal = strVal & UniConvNum(.vspdData2.Text,0) & parent.gRowSep
		        
		        ReDim Preserve TmpBufferVal(iValCnt)
		        TmpBufferVal(iValCnt) = strVal
		        iValCnt = iValCnt + 1
		        lGrpCnt = lGrpCnt + 1
		        
		    Case ggoSpread.DeleteFlag
				strDel = ""
				strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep				'⊙: D=Delete
		        .vspdData2.Col = C_ResourceCd						'2
		        strDel = strDel & UCase(Trim(.vspdData2.Text)) & parent.gColSep
		        .vspdData2.Col = C_Rank								'3
		        strDel = strDel & UniConvNum(.vspdData2.Text,0) & parent.gRowSep
		        
		        ReDim Preserve TmpBufferDel(iDelCnt)
		        TmpBufferDel(iDelCnt) = strDel
		        iDelCnt = iDelCnt + 1
				lGrpcnt = lGrpcnt + 1
				             
		End Select
    Next
	
	End With
	
	iTotalStrDel = Join(TmpBufferDel, "")
	iTotalStrVal = Join(TmpBufferVal, "")
	
	frm1.txtMaxRows.value = lGrpCnt - 1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value = iTotalStrDel & iTotalStrVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	
    DbSave = True    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgStrPrevKey2 = ""
	
	frm1.vspdData2.MaxRows = 0 
	
	Call DisableToolBar(parent.TBC_QUERY)  
	If DbDtlQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If			
	
	Call SetToolbar("1100111100111111")

	lgIntFlgMode = parent.OPMD_UMODE
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
Function SheetFocus(lRow, lCol)
	frm1.vspdData2.focus
	frm1.vspdData2.Row = lRow
	frm1.vspdData2.Col = lCol
	frm1.vspdData2.Action = 0
	frm1.vspdData2.SelStart = 0
	frm1.vspdData2.SelLength = len(frm1.vspdData.Text)
End Function

'==============================================================================
' Function : ConvToSec()
' Description : 저장시에 각 시간 데이터들을 초로 환산 
'==============================================================================
Function ConvToSec(Str)
	If Str = "" Then
		ConvToSec = 0
	ElseIf Len(Trim(Str)) = 8 Then
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
	Else
		ConvToSec = -999999
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자원구성정보등록</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
			 					<TR>
			 						<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>
			 						<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>라우팅</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=15 MAXLENGTH=7 tag="12XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConRouting()">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutNm" SIZE=20 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=2 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE Class="TB3" WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS="TD5" NOWRAP>주 라우팅</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" NAME="rdoMajorRouting" ID="rdoMajorRouting1" Value="Y" CLASS="RADIO" tag="24X" CHECKED><LABEL FOR="rdoMajorRouting1">예</LABEL>
													   <INPUT TYPE="RADIO" NAME="rdoMajorRouting" ID="rdoMajorRouting2" Value="N" CLASS="RADIO" tag="24X"><LABEL op="rdoMajorRouting2">아니오</LABEL></TD>
								<TD CLASS="TD5" NOWRAP>유효기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME SIZE="10" MAXLENGTH="10" ALT="시작일" tag="24X1"> </OBJECT>');</SCRIPT>								
									&nbsp;~&nbsp; 
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24X1" ALT="종료일" MAXLENGTH="10" SIZE="10"> </OBJECT>');</SCRIPT>								
								</TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" colspan=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 id="A" width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 id="B" WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOprNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBomNo" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
