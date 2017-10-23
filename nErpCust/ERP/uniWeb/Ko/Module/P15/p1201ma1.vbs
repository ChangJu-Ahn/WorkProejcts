'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p1201mb1.asp"								'☆: 비지니스 로직 ASP명 

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p1201mb2.asp"								'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "p1201mb3.asp"								'☆: 비지니스 로직 ASP명 

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
Dim C_HdnInsideFlg

' Grid 2(vspdData2) - Operation 
Dim C_Select
Dim C_ChildItemCd
Dim C_ChildItemNm
Dim C_ChildItemSpec
Dim C_IssuedSlCd
Dim C_IssuedSlNm
Dim C_IssuedUnit
Dim C_PrntItemCd
Dim C_PrntItemNm
Dim C_PrntItemSpec
Dim C_ChildItemSeq
Dim C_ValidFromDt1
Dim C_ValidToDt1
Dim C_HiddenFlg	

Dim lgIntPrevKey

Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2

Dim lgButtonSelection

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
		C_FixRunTime            = 10
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
		C_HdnInsideFlg			= 25
	End If

	If pvGridId = "*" Or pvGridId = "B" Then
		' Grid 2(vspdData2) - Operation 
		C_Select		= 1
		C_ChildItemCd	= 2
		C_ChildItemNm	= 3
		C_ChildItemSpec	= 4
		C_IssuedSlCd	= 5
		C_IssuedSlNm	= 6
		C_IssuedUnit	= 7
		C_PrntItemCd	= 8
		C_PrntItemNm	= 9
		C_PrntItemSpec	= 10
		C_ChildItemSeq	= 11
		C_ValidFromDt1	= 12
		C_ValidToDt1	= 13
		C_HiddenFlg		= 14
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
    
    lgStrPrevKey = ""							'initializes Previous Key
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
    
    lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
    
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
		'frm1.txtOprNo.Value		= ReadCookie("txtOprNo")
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtRoutingNo", ""
	WriteCookie "txtRoutingNm", ""
	'WriteCookie "txtOprNo", ""
	frm1.txtBaseDt.Text = StartDate
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
			
			.MaxCols = C_HdnInsideFlg + 1
			.MaxRows = 0

			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_OprNo,				"공정", 7,,,3,2
			ggoSpread.SSSetEdit		C_WCCd,					"작업장", 10,,,7,2
			ggoSpread.SSSetEdit		C_JobCd,				"공정작업", 10
			ggoSpread.SSSetEdit		C_JobNm,				"공정작업명", 12
			ggoSpread.SSSetEdit		C_InsideFlg,			"공정타입", 10
			ggoSpread.SSSetEdit		C_MfgLt,				"제조L/T", 7,1,,3
			ggoSpread.SSSetTime		C_QueueTime,			"Queue시간", 10, 2,1 ,1
			ggoSpread.SSSetTime		C_SetupTime,			"설치시간", 10, 2,1 ,1
			ggoSpread.SSSetTime		C_WaitTime,				"대기시간", 10, 2,1 ,1
			ggoSpread.SSSetTime		C_FixRunTime,			"고정가동시간", 10, 2,1 ,1
			ggoSpread.SSSetTime		C_RunTime,				"변동가동시간", 10, 2,1 ,1
			ggoSpread.SSSetFloat	C_ItemQtyForRunTime,	"기준수량", 15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_UnitOfItemQtyForRunTime, "기준단위", 10,,,3,2
			ggoSpread.SSSetTime		C_MoveTime,				"이동시간", 10, 2,1 ,1
			ggoSpread.SSSetEdit		C_OverLapOpr,			"Overlap 공정", 7,,,3,2
			ggoSpread.SSSetEdit		C_OverLapLt,			"Overlap L/T",8,1
			ggoSpread.SSSetEdit		C_BpCd,					"외주처", 10,,,18,2
			ggoSpread.SSSetEdit		C_CurCd,				"통화", 6,,,3,2
			'ggoSpread.SSSetFloat	C_UnitPriceOfOprSubcon,	"공정외주단가",15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_UnitPriceOfOprSubcon,	"공정외주단가",15,"C"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_TaxType,				"VAT유형", 16,,,20,2
			ggoSpread.SSSetEdit		C_MilestoneFlg,			"Milestone", 10
			ggoSpread.SSSetEdit		C_RoutOrder,			"공정단계", 10
			ggoSpread.SSSetDate 	C_ValidFromDt,			"시작일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ValidToDt,			"종료일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_HdnInsideFlg,			"공정타입", 10

			Call ggoSpread.SSSetColHidden(C_HdnInsideFlg, C_HdnInsideFlg, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
			ggoSpread.SSSetSplit2(2)										'frozen 기능추가 
			.ReDraw = True
    
		End With
	End If
	
	If pvGridId = "*" Or pvGridId = "B" Then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	
		With frm1.vspdData2
	
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20030109",, parent.gAllowDragDropSpread

			.ReDraw = False

			.MaxCols = C_HiddenFlg + 1
			.MaxRows = 0

			Call GetSpreadColumnPos("B")
	
			ggoSpread.SSSetCheck	C_Select ,		"",				2,,,1
			ggoSpread.SSSetEdit	C_ChildItemCd,	"자품목", 15 
			ggoSpread.SSSetEdit	C_ChildItemNm,	"자품목명",		18
			ggoSpread.SSSetEdit	C_ChildItemSpec, "자품목규격",	18
			ggoSpread.SSSetEdit	C_IssuedSlCd,	"출고창고",		10 
			ggoSpread.SSSetEdit	C_IssuedSlNm,	"출고창고명",	18 
			ggoSpread.SSSetEdit	C_IssuedUnit,	"출고단위",		8  
			ggoSpread.SSSetEdit	C_PrntItemCd,	"모품목",		15  
			ggoSpread.SSSetEdit	C_PrntItemNm,	"모품목명",		18
			ggoSpread.SSSetEdit	C_PrntItemSpec, "모품목규격",	18
			ggoSpread.SSSetEdit	C_ChildItemSeq, "순서",			6,	1
			ggoSpread.SSSetDate C_ValidFromDt1,	"시작일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate C_ValidToDt1,	"종료일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit	C_HiddenFlg,	"할당구분",		3

			Call ggoSpread.SSSetColHidden(C_HiddenFlg, C_HiddenFlg, True)
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
			ggoSpread.SSSetSplit(2)										'frozen 기능추가 
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
			ggoSpread.SpreadLock 2, -1, .vspdData2.MaxCols
			.vspdData2.ReDraw = True
		End If
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
			C_FixRunTime            = iCurColumnPos(10)  
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
			C_HdnInsideFlg			= iCurColumnPos(25)
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Select		= iCurColumnPos(1)
			C_ChildItemCd	= iCurColumnPos(2)
			C_ChildItemNm	= iCurColumnPos(3)
			C_ChildItemSpec	= iCurColumnPos(4)
			C_IssuedSlCd	= iCurColumnPos(5)
			C_IssuedSlNm	= iCurColumnPos(6)
			C_IssuedUnit	= iCurColumnPos(7)
			C_PrntItemCd	= iCurColumnPos(8)
			C_PrntItemNm	= iCurColumnPos(9)
			C_PrntItemSpec	= iCurColumnPos(10)
			C_ChildItemSeq	= iCurColumnPos(11)
			C_ValidFromDt1	= iCurColumnPos(12)
			C_ValidToDt1	= iCurColumnPos(13)
			C_HiddenFlg		= iCurColumnPos(14)
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
	Call ggoSpread.ReOrderingSpreadData()
	If gActiveSpdSheet.id = "B" Then
		Call DbDtlQueryOk(1)
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
		Call DisplayMsgBox("971012", "X", "공장", "X")
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

	arrParam(0) = "라우팅 팝업"	
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

'------------------------------------------  SetRouting()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRouting(byval arrRet)
	frm1.txtRoutNo.Value    = arrRet(0)
	frm1.txtRoutNm.Value    = arrRet(1)
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
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

Function btnAutoSel_onClick()

	Dim iRows
	Dim iRow	
	
	frm1.vspdData2.ReDraw = false
	
	with frm1.vspdData2	
		iRows = .maxRows			
		for iRow=1 to iRows
			.Col = C_Select
			.Row = iRow
			If lgButtonSelection = "SELECT" Then
				If .value = 1 Then
					.value = 0
					Call vspdData2_ButtonClicked(C_Select, iRow, 0)
				End If	
			Else
				If .value = 0 Then
					.value = 1
					Call vspdData2_ButtonClicked(C_Select, iRow, 1)
				End If	
				
			End If	
		next 		
	end with	
	
	frm1.vspdData2.ReDraw = true

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "전체선택"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "전체선택취소"
	End If

End Function

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	Dim IntRetCD

	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData1
    Call SetPopupMenuItemInf("0000111111")
	
	If frm1.vspdData1.MaxRows <= 0 Or Col < 0 Or Row <= 0 Then
		Exit Sub
	End If
	
	
	If lgOldRow <> Row Then
		
		If ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
			If IntRetCD = vbNo Then
				Exit Sub
			End If
		End If
		
		frm1.vspdData1.Row = Row
		frm1.vspdData1.Col = C_OprNo
		frm1.hOprNo.value = Trim(frm1.vspdData1.Text)
		
		frm1.vspdData1.Col = C_HdnInsideFlg
		frm1.hInsideFlg.value = UCase(Trim(frm1.vspdData1.Text))
		
		lgOldRow = Row
		
		frm1.vspdData2.MaxRows = 0
		
		LayerShowHide(1)
		
		If DbDtlQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If				
		
	End If
	
End Sub

Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If frm1.vspdData1.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If
	
	Call vspdData1_Click(NewCol, NewRow)

End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
    Call SetPopupMenuItemInf("0000111111")

	If frm1.vspdData2.MaxRows <= 0 Or Col < 0 Then
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
	If NewCol = C_Select Or Col = C_Select Then
		Cancel = True
		Exit Sub
	End If
	
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button, Shift, X, Y)
		
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : Check button clicked
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	ggoSpread.Source = frm1.vspdData2
    With frm1.vspdData2
		.Row = Row
		.Col = C_HiddenFlg
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
				lgLngCnt = lgLngCnt + 1
			Else
				If lgAfterQryFlg = True Then
					ggoSpread.SSDeleteFlag Row,Row
					lgLngCnt = lgLngCnt - 1
				End If
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
				lgLngCnt = lgLngCnt + 1
			Else
				If lgAfterQryFlg = True Then
					ggoSpread.SSDeleteFlag Row,Row
					lgLngCnt = lgLngCnt - 1
				End If
			End If			
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
		If lgIntPrevKey <> 0 Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			LayerShowHide(1)
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbDtlQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
'			Call DbDtlQuery
		End If     
    End if
    
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
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtRoutNo.value = "" Then
		frm1.txtRoutNm.value = ""
	End If
	
    
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
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
    End If     																'☜: Query db data
       
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
    If ggoSpread.SSDefaultCheck = False Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     							                                                  '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData2	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
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
    Call parent.FncExport(parent.C_SINGLEMULTI)                                                   '☜: Protect system from crashing
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

	Call SetToolBar("11001000000111")
				
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		frm1.vspdData1.Col = C_OprNo
		frm1.vspdData1.Row = 1
	
		frm1.hOprNo.value = Trim(frm1.vspdData1.Text) 
	
		frm1.vspdData1.Col = C_HdnInsideFlg
		frm1.vspdData1.Row = 1
	
		frm1.hInsideFlg.value = Trim(frm1.vspdData1.Text) 
	
		lgOldRow = 1
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		Call DisableToolBar(parent.TBC_QUERY)  
		If DbDtlQuery = False Then
			Call RestoreToolBar()
			Exit Function
		End If
	Else
		Call LayerShowHide(0)
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 
    Dim strVal
    
    DbDtlQuery = False
    
    'LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "전체선택"
        
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.value)
		strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.value)
		strVal = strVal & "&txtBomNo=" & Trim(.txtBomNo.value)
		strVal = strVal & "&txtBaseDt=" & Trim(frm1.hBaseDt.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
    Else
		strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtRoutNo=" & Trim(.txtRoutNo.value)
		strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.value)
		strVal = strVal & "&txtBomNo=" & Trim(.txtBomNo.value)
		strVal = strVal & "&txtBaseDt=" & Trim(frm1.txtBaseDt.text)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)														'☆: 조회 성공후 실행로직 
	Dim i	
    '-----------------------
    'Reset variables area
    '-----------------------
    
    frm1.vspdData2.redraw = false
    With frm1.vspdData2
		For	 i = LngMaxRow To .MaxRows
			.Row = i
			.Col = C_HiddenFlg
			If .Text = "Y" Then
				.Col = C_Select
				.Value = 1
			End If 	  
			
			If frm1.hInsideFlg.value = "N" Then
				ggoSpread.SpreadLock C_Select, i, C_Select
			End If
		Next		
	End With
	frm1.vspdData2.Redraw = True
	
	lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	frm1.btnAutoSel.disabled = False
	
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()
    Dim lRow, lGrpCnt        
    Dim strVal
    Dim strDel
	Dim ITemp
	Dim TmpBufferVal, TmpBufferDel
	Dim iTotalStrVal, iTotalStrDel
	Dim iValCnt, iDelCnt

    DbSave = False                                                          
    
    LayerShowHide(1)
		
    'On Error Resume Next                                                   

	With frm1
		.txtMode.value = parent.UID_M0002
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    
    lGrpCnt = 0
    ReDim TmpBufferVal(0) : ReDim TmpBufferDel(0)
    iValCnt = 0 : iDelCnt = 0
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    For lRow = 1 To .vspdData2.MaxRows
    
		ggoSpread.Source = .vspdData2 
    
        .vspdData2.Row = lRow
        .vspdData2.Col = 0
        
        ITemp = ""
        
        Select Case .vspdData2.Text
                
            Case ggoSpread.UpdateFlag
				.vspdData2.Col = C_Select

				If .vspdData2.Value = 1 Then
					ITemp = "Y"
				Else
					ITemp = "N"
				End If 											'☜: 신규 
				
				.vspdData2.Col = C_HiddenFlg

				If ITemp = "Y" And .vspdData2.Text = "N" Then
					
					strVal = ""
					
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep

					strVal = strVal & Trim(.txtBomNo.value) & parent.gColSep						'Prnt BOM No Data

					.vspdData2.Col = C_ChildItemSeq	'10
					strVal = strVal & Trim(.vspdData2.Text) & parent.gColSep
                
					.vspdData2.Col = C_ChildItemCd	'3
					strVal = strVal & Trim(.vspdData2.Text) & parent.gRowSep
					
					ReDim Preserve TmpBufferVal(iValCnt)
					
					TmpBufferVal(iValCnt) = strVal
					
					iValCnt = iValCnt + 1
					
					lGrpCnt = lGrpCnt + 1
					
				ElseIf ITemp = "N" And .vspdData2.Text = "Y" Then
				
					strDel = ""
					
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep

					strDel = strDel & Trim(.txtBomNo.value) & parent.gColSep						'Prnt BOM No Data
					
					.vspdData2.Col = C_ChildItemSeq	'10
					strDel = strDel & Trim(.vspdData2.Text) & parent.gColSep
                
					.vspdData2.Col = C_ChildItemCd	'3
					strDel = strDel & Trim(.vspdData2.Text) & parent.gRowSep
					
					ReDim Preserve TmpBufferDel(iDelCnt)
					
					TmpBufferDel(iDelCnt) =  strDel
					
					iDelCnt = iDelCnt + 1
				
					lGrpCnt = lGrpCnt + 1
					
				End If 											'☜: 신규 
        End Select
    Next

	If lGrpCnt > 0 Then
		iTotalStrDel = Join(TmpBufferDel, "")
		iTotalStrVal = Join(TmpBufferVal, "")
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = iTotalStrDel & iTotalStrVal
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	Else
		Call LayerShowHide(0)									'☜: 비지니스 ASP 를 가동 
	End If
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0

    frm1.vspdData2.MaxRows = 0
	
	If DbDtlQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If     
	
	lgIntFlgMode = parent.OPMD_UMODE
	
End Function