Const BIZ_PGM_QRY_ID						= "p1203mb1.asp"					'☆: Detail Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID						= "p1203mb2.asp"					'☆: Save 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID						= "p1203mb3.asp"					'☆: Delete 비지니스 로직 ASP명 
Const BIZ_PGM_COPY_ID						= "p1203mb4.asp"					'☆: 표준라우팅 복사 비지니스 로직 ASP명 
Const BIZ_PGM_LOOKUPWC_ID					= "p1203mb7.asp"
Const BIZ_PGM_JUMPCOMPALLOC_ID				= "p1201ma1"

Dim C_OprNo
Dim C_WCCd
Dim C_WCPopup
Dim C_JobCd
Dim C_JobNm
Dim C_InsideFlg
Dim C_MfgLt
Dim C_QueueTime
Dim C_SetupTime
Dim C_WaitTime
Dim C_FixRunTime
Dim C_RunTime
Dim C_RunTimeQty
Dim C_RunTimeUnit
Dim C_UnitPopup
Dim C_MoveTime
Dim C_OverLapOpr
Dim C_OverLapLt
Dim C_BpCd
Dim C_BpPopup
Dim C_BpNm
Dim C_CurCd
Dim C_CurPopup
Dim C_SubconPrc
Dim C_TaxType
Dim C_TaxPopup
Dim C_MilestoneFlg
Dim C_InspFlg
Dim C_RoutOrder
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_HiddenInsideFlg
Dim C_HiddenRoutOrder	

Dim lgStrPrevKey2			  'Routing Copy 용 이전 Key 값	
Dim IsOpenPop					 'Popup
Dim lgRdoOldVal
Dim lgLastOpNo, lgLastOpRowNo
Dim lgItemBaseUnit			 'Base Unit for Inserting Rows

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_OprNo			= 1
	C_WCCd			= 2
	C_WCPopup		= 3
	C_JobCd			= 4
	C_JobNm			= 5
	C_InsideFlg		= 6
	C_MfgLt			= 7
	C_QueueTime		= 8
	C_SetupTime		= 9
	C_WaitTime		= 10
	C_FixRunTime	= 11
	C_RunTime		= 12
	C_RunTimeQty	= 13
	C_RunTimeUnit	= 14
	C_UnitPopup		= 15
	C_MoveTime		= 16
	C_OverLapOpr	= 17
	C_OverLapLt		= 18
	C_BpCd			= 19
	C_BpPopup		= 20
	C_BpNm			= 21
	C_CurCd			= 22
	C_CurPopup		= 23
	C_SubconPrc		= 24
	C_TaxType		= 25
	C_TaxPopup		= 26
	C_MilestoneFlg	= 27
	C_InspFlg		= 28
	C_RoutOrder		= 29
	C_ValidFromDt	= 30
	C_ValidToDt		= 31
	C_HiddenInsideFlg = 32
	C_HiddenRoutOrder = 33
End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                                       '⊙: initializes sort direction
	lgItemBaseUnit = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtValidFromDt.Text	= StartDate
	frm1.txtValidToDt.Text		= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	frm1.rdoMajorRouting1.checked = True
	
	lgRdoOldVal = 1	
End Sub

Sub ReadCookVal()
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value	= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value	= ReadCookie("txtPlantNm")
		frm1.txtItemCd.Value	= ReadCookie("txtItemCd")
		frm1.txtItemNm.value	= ReadCookie("txtItemNm")
		frm1.txtRoutingNo.Value	= ReadCookie("txtRoutingNo")
		frm1.txtRoutingNm.value	= ReadCookie("txtRoutingNm")
		Call txtPlantCd_OnChange()
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtRoutingNo", ""
	WriteCookie "txtRoutingNm", ""	
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
        
	Call initSpreadPosVariables()
    
    With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_HiddenRoutOrder + 1
		.MaxRows = 0
	
		Call AppendNumberPlace("6","3","0")
	
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_OprNo, "공정", 7,,,3,2
		ggoSpread.SSSetEdit		C_WCCd, "작업장", 10,,,7,2
		ggoSpread.SSSetButton 	C_WCPopup
		ggoSpread.SSSetCombo	C_JobCd, "공정작업", 10
		ggoSpread.SSSetCombo	C_JobNm, "공정작업명", 15
		ggoSpread.SSSetEdit		C_InsideFlg, "공정타입", 10
		ggoSpread.SSSetFloat	C_MfgLt,	"제조L/T",10,"6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetTime		C_QueueTime, "Queue시간", 10, 2, 1 ,1
		ggoSpread.SSSetTime		C_SetupTime, "설치시간", 10, 2, 1 ,1
		ggoSpread.SSSetTime		C_WaitTime, "대기시간", 10, 2, 1 ,1
		ggoSpread.SSSetTime		C_FixRunTime, "고정가동시간", 10, 2, 1 ,1
		ggoSpread.SSSetTime		C_RunTime, "변동가동시간", 10, 2, 1 ,1
		ggoSpread.SSSetFloat	C_RunTimeQty,"기준수량",15, parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"P"
		ggoSpread.SSSetEdit 	C_RunTimeUnit, "기준단위", 10,,,3,2
		ggoSpread.SSSetButton 	C_UnitPopup
		ggoSpread.SSSetTime		C_MoveTime, "이동시간", 10, 2,1 ,1
		ggoSpread.SSSetEdit		C_OverLapOpr, "Overlap 공정", 11,,,3,2
		ggoSpread.SSSetFloat	C_OverLapLt,"Overlap L/T",12,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,1,FALSE,"Z" 
		ggoSpread.SSSetEdit		C_BpCd, "외주처", 10,,,18,2
		ggoSpread.SSSetButton 	C_BpPopup
		ggoSpread.SSSetEdit		C_BpNm, "외주처명", 20
		ggoSpread.SSSetEdit		C_CurCd, "통화", 6,,,3,2
		ggoSpread.SSSetButton 	C_CurPopup
		'ggoSpread.SSSetFloat	C_SubconPrc,"공정외주단가", 15, parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_SubconPrc,"공정외주단가", 15, "C", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_TaxType, "VAT유형", 8,,,5,2
		ggoSpread.SSSetButton	C_TaxPopup
		ggoSpread.SSSetCombo	C_MilestoneFlg, "Milestone", 10
		ggoSpread.SSSetCombo	C_InspFlg,	"검사여부", 10
		ggoSpread.SSSetEdit		C_RoutOrder, "공정단계", 10
		ggoSpread.SSSetDate 	C_ValidFromDt, "시작일", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt, "종료일", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_HiddenInsideFlg, "공정타입", 10
		ggoSpread.SSSetEdit		C_HiddenRoutOrder, "공정단계", 10
	
		Call ggoSpread.MakePairsColumn(C_WCCd, C_WCPopup)
		Call ggoSpread.MakePairsColumn(C_RunTimeUnit, C_UnitPopup)
		Call ggoSpread.MakePairsColumn(C_BpCd, C_BpPopup)
		Call ggoSpread.MakePairsColumn(C_CurCd, C_CurPopup)
		Call ggoSpread.MakePairsColumn(C_TaxType, C_TaxPopup)

		Call ggoSpread.SSSetColHidden(C_HiddenInsideFlg, C_HiddenRoutOrder, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

		.ReDraw = True
	End With
    
	ggoSpread.SSSetSplit2(3)										'frozen 기능추가 
    
	Call SetSpreadLock()
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_OprNo,		-1, C_OprNo
		ggoSpread.SpreadLock C_InsideFlg,	-1, C_InsideFlg
		ggoSpread.SpreadLock C_BpCd,		-1, C_BpPopup
		ggoSpread.SpreadLock C_BpNm,		-1, C_BpNm
		ggoSpread.SpreadLock C_CurCd,		-1, C_CurPopup
		ggoSpread.SpreadLock C_SubconPrc,	-1, C_SubconPrc
		ggoSpread.SpreadLock C_TaxType,		-1, C_TaxPopup
		ggoSpread.SpreadLock C_RoutOrder,	-1, C_RoutOrder
		ggoSpread.SpreadLock C_ValidFromDt,	-1, C_ValidFromDt
	
		ggoSpread.SSSetRequired C_WCCd, 		-1, -1
		ggoSpread.SSSetRequired	C_MilestoneFlg,	-1, -1
		ggoSpread.SSSetRequired	C_InspFlg,		-1, -1
		ggoSpread.SSSetRequired C_ValidToDt, 	-1, -1
		ggoSpread.SSSetRequired C_RunTimeUnit, 	-1, -1
		
		ggoSpread.SSSetProtected .vspdData.MaxCols, -1
		.vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal InOutType)
	ggoSpread.Source = frm1.vspdData
    If InOutType = "N" Then
		ggoSpread.SSSetRequired 	C_OprNo, 		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_WCCd, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_InsideFlg, 	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_BpNm, 		pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected 	C_RoutOrder, 	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_MilestoneFlg,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_InspFlg,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ValidFromDt, 	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ValidToDt, 	pvStartRow, pvEndRow

		ggoSpread.SpreadUnLock		C_BpCd,			pvStartRow, C_BpPopup, pvEndRow
		ggoSpread.SpreadUnLock		C_CurCd,		pvStartRow, C_CurPopup, pvEndRow
		ggoSpread.SpreadUnLock		C_SubconPrc,	pvStartRow, C_SubconPrc, pvEndRow
		ggoSpread.SpreadUnLock		C_TaxType,		pvStartRow, C_TaxPopup, pvEndRow

		ggoSpread.SSSetRequired		C_BpCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_CurCd,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_SubconPrc,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_TaxType,		pvStartRow, pvEndRow

	Else
		ggoSpread.SSSetRequired 	C_OprNo, 		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_WCCd, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_InsideFlg, 	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_BpCd, 		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_BpPopup, 		pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected 	C_BpNm, 		pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected 	C_CurCd, 		pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected 	C_CurPopup,		pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected 	C_SubconPrc,	pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_TaxType,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected 	C_TaxPopup,		pvStartRow, pvEndRow    
		ggoSpread.SSSetProtected 	C_RoutOrder, 	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_MilestoneFlg,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_InspFlg,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ValidFromDt, 	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ValidToDt, 	pvStartRow, pvEndRow
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
			C_OprNo			= iCurColumnPos(1)
			C_WCCd			= iCurColumnPos(2)
			C_WCPopup		= iCurColumnPos(3)
			C_JobCd			= iCurColumnPos(4)
			C_JobNm			= iCurColumnPos(5)
			C_InsideFlg		= iCurColumnPos(6)
			C_MfgLt			= iCurColumnPos(7)
			C_QueueTime		= iCurColumnPos(8)
			C_SetupTime		= iCurColumnPos(9)
			C_WaitTime		= iCurColumnPos(10)
			C_FixRunTime	= iCurColumnPos(11)
			C_RunTime		= iCurColumnPos(12)
			C_RunTimeQty	= iCurColumnPos(13)
			C_RunTimeUnit	= iCurColumnPos(14)
			C_UnitPopup		= iCurColumnPos(15)
			C_MoveTime		= iCurColumnPos(16)
			C_OverLapOpr	= iCurColumnPos(17)
			C_OverLapLt		= iCurColumnPos(18)
			C_BpCd			= iCurColumnPos(19)
			C_BpPopup		= iCurColumnPos(20)
			C_BpNm			= iCurColumnPos(21)
			C_CurCd			= iCurColumnPos(22)
			C_CurPopup		= iCurColumnPos(23)
			C_SubconPrc		= iCurColumnPos(24)
			C_TaxType		= iCurColumnPos(25)
			C_TaxPopup		= iCurColumnPos(26)
			C_MilestoneFlg	= iCurColumnPos(27)
			C_InspFlg		= iCurColumnPos(28)
			C_RoutOrder		= iCurColumnPos(29)
			C_ValidFromDt	= iCurColumnPos(30)
			C_ValidToDt		= iCurColumnPos(31)
			C_HiddenInsideFlg = iCurColumnPos(32)
			C_HiddenRoutOrder = iCurColumnPos(33)
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
	Dim iIntCnt
	
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()

    frm1.vspdData.redraw = False
    Call InitSpreadSheet()
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)

    For iIntCnt = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = iIntCnt
		frm1.vspdData.Col = C_HiddenInsideFlg
	
		If UCase(Trim(frm1.vspdData.Text)) = "N" Then
			Call SetFieldProp(iIntCnt, "N")
		Else
			Call SetFieldProp(iIntCnt, "Y")
		End IF
		
    Next

	Call ProtectMilestone(1)
	
	frm1.vspdData.Redraw = True
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

    Dim strCboCd 
    Dim strCboNm 
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    strCboCd = "" & vbTab
    strCboNm = "" & vbTab
  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    ggoSpread.Source = frm1.vspdData
    lgF0 = "" & Chr(11) & lgF0
    lgF1 = "" & Chr(11) & lgF1
    ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_JobCd
    ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_JobNm
  
    '****************************
    'MileStone Flag Setting
    '****************************

    strCboCd = ""
    strCboCd = "Y" & vbTab & "N"
    
    ggoSpread.SetCombo strCboCd, C_MilestoneFlg
    
    '****************************
    'Insp Flag Setting
    '****************************

    strCboCd = ""
    strCboCd = "Y" & vbTab & "N"
    
    ggoSpread.SetCombo strCboCd,C_InspFlg
  
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		'.ReDraw = False
		
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.Col = C_JobCd
			intIndex = .value
			.col = C_JobNm
			.value = intindex
			
			.Col = C_HiddenInsideFlg 
		Next	
		'.ReDraw = True
	End With
End Sub

Function SetFieldProp(ByVal lRow, ByVal sType)
	ggoSpread.Source = frm1.vspdData
	
	If sType = "N" Then			'외주 공정이면 
		ggoSpread.SpreadUnLock	C_BpCd,		lRow, C_BpPopup, lRow
		ggoSpread.SpreadUnLock	C_CurCd,	lRow, C_CurPopup, lRow
		ggoSpread.SpreadUnLock	C_SubconPrc,lRow, C_SubconPrc, lRow
		ggoSpread.SpreadUnLock	C_TaxType,	lRow, C_TaxPopup, lRow

		ggoSpread.SSSetRequired	C_BpCd,			lRow, lRow
		ggoSpread.SSSetRequired	C_CurCd,		lRow, lRow
		ggoSpread.SSSetRequired	C_SubconPrc,	lRow, lRow
		ggoSpread.SSSetRequired	C_TaxType,		lRow, lRow

	ElseIf sType = "Y" Then		'사내 공정이면 
		ggoSpread.SpreadLock	C_BpCd,			lRow, C_BpPopup,  lRow
		ggoSpread.SpreadLock	C_CurCd,		lRow, C_CurPopup, lRow
		ggoSpread.SpreadLock	C_SubconPrc,	lRow, C_SubconPrc, lRow
		ggoSpread.SpreadLock	C_TaxType,		lRow, C_TaxPopup, lRow

		ggoSpread.SSSetProtected	C_BpCd,			lRow, lRow
		ggoSpread.SSSetProtected	C_CurCd,		lRow, lRow
		ggoSpread.SSSetProtected	C_SubconPrc,	lRow, lRow
		ggoSpread.SSSetProtected	C_TaxType,		lRow, lRow
	End If
	
End Function

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtBomNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If frm1.txtItemCd1.value= "" Then
		Call DisplayMsgBox("971012", "X", "품목", "X")
		frm1.txtItemCd1.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
		
	IsOpenPop = True

	arrParam(0) = "BOM팝업"							' 팝업 명칭 
	arrParam(1) = "P_BOM_HEADER"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBomNo.Value)				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				  " AND ITEM_CD = "	& FilterVar(frm1.txtItemCd.value, "''", "S")
				  
				  										' Where Condition
	arrParam(5) = "BOM"								' TextBox 명칭 
	
    arrField(0) = "BOM_NO"								' Field명(0)
    arrField(1) = "DESCRIPTION"							' Field명(1)
    
    arrHeader(0) = "BOM Type"						' Header명(0)
    arrHeader(1) = "BOM설명"						' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBomNo.focus
	
End Function

'===========================================================================
' Function Name : OpenRoutingNo
' Function Desc : OpenRoutingNo Reference Popup
'===========================================================================
Function OpenRoutingNo()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtRoutingNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012", "X", "품목", "X")
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True
		
	arrParam(0) = "라우팅 팝업"					' 팝업 명칭 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtRoutingNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				  " AND ITEM_CD = "	& FilterVar(frm1.txtItemCd.value, "''", "S")
	arrParam(5) = "라우팅"			
	
    arrField(0) = "ROUT_NO"							' Field명(0)
    arrField(1) = "DESCRIPTION"						' Field명(1)
    arrField(2) = "MAJOR_FLG"						' Field명(1)
    
    arrHeader(0) = "라우팅"						' Header명(0)
    arrHeader(1) = "라우팅명"					' Header명(1)
    arrHeader(2) = "주라우팅여부"				' Header명(1)
    
    arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetRoutingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtRoutingNo.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode,ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If iPos = 1 Then
		If UCase(frm1.txtItemCd1.className) = UCase(parent.UCN_PROTECTED) Then 
			Exit Function
		End If
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X" , "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = "12!MO"							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분: From To를 입력할 것 
	arrParam(3) = ""							' Default Value

	If iPos = 1 Then
		arrParam(4) = BaseDate
	End If	
	
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
		Call SetItemInfo(arrRet,iPos)
	End If	
	
	Call SetFocusToDocument("M")
	If iPos = "0" Then
		frm1.txtItemCd.focus 
	Else
		frm1.txtItemCd1.focus
	End If

End Function


'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenWcPopup()
'	Description : WcPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWcPopup(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " AND VALID_TO_DT >=  " & FilterVar(BaseDate , "''", "S") & ""
	arrParam(5) = "작업장"	
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "HH" & parent.gcolsep & "INSIDE_FLG"
    arrField(3) = "CASE WHEN INSIDE_FLG=" & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("사내", "''", "S") & " ELSE " & FilterVar("외주", "''", "S") & " END"
    arrField(4) = "dbo.ufn_GetCodeName(" & FilterVar("P1013", "''", "S") & ", WC_MGR)"	
    
    arrHeader(0) = "작업장"	
    arrHeader(1) = "작업장명"
    arrHeader(2) = "작업장구분"
    arrHeader(3) = "작업장구분"
    arrHeader(4) = "작업장담당자"
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWc(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_WcCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenWcPopup()
'	Description : WcPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit(ByVal str)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & " "			
	arrParam(5) = "단위"			
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
   
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"		
    
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_RunTimeUnit,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
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

	arrParam(0) = "외주처팝업"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "외주처"			
	
    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"
    arrField(2) = "ED15" & parent.gcolsep & "BP_TYPE"
    arrField(3) = "ED15" & parent.gcolsep & "CURRENCY"
    arrField(4) = "ED15" & parent.gcolsep & "VAT_TYPE"
        
    arrHeader(0) = "BP"
    arrHeader(1) = "BP명"
    arrHeader(2) = "Bp 구분"
    arrHeader(3) = "통화"
    arrHeader(4) = "VAT유형"
        
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetBizPartner(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_BpCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
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

	arrParam(0) = "통화팝업"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "통화"			
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
    
    arrHeader(0) = "통화"		
    arrHeader(1) = "통화명"		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCurrency(arrRet)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_CurCd,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenRoutCopy()  -------------------------------------------------
'	Name : OpenRoutCopy()
'	Description : Routing Copy Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRoutCopy()
	Dim arrRet
	
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then		
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "표준라우팅 복사"						' 팝업 명칭 
	arrParam(1) = "(SELECT Distinct ROUT_NO, DESCRIPTION FROM P_STANDARD_ROUTING WHERE PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ") A"	' TABLE 명칭 
	arrParam(2) = ""										' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = ""	
	arrParam(5) = "표준라우팅"							' TextBox 명칭 
	
    arrField(0) = "A.ROUT_NO"								' Field명(0)
    arrField(1) = "A.DESCRIPTION"							' Field명(0)
    
    arrHeader(0) = "표준라우팅"								' Header명(0)
    arrHeader(1) = "표준라우팅명"							' Header명(0)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call  SetRoutCopy(arrRet)
	End If	
	
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

	arrParam(0) = "VAT유형"						
	arrParam(1) = "B_MINOR, B_CONFIGURATION"						
	
	arrParam(2) = Trim(frm1.vspdData.text)	
		
	arrParam(4) = "B_MINOR.MAJOR_CD=" & FilterVar("B9001", "''", "S") & " AND B_MINOR.MINOR_CD=B_CONFIGURATION.MINOR_CD "							
	arrParam(4) = arrParam(4) & "AND B_MINOR.MAJOR_CD=B_CONFIGURATION.MAJOR_CD AND B_CONFIGURATION.SEQ_NO=1"
	arrParam(5) = "VAT유형"							
	
    arrField(0) = "B_MINOR.MINOR_CD"					
    arrField(1) = "B_MINOR.MINOR_NM"
    arrField(2) = "F5" & parent.gColSep & "B_CONFIGURATION.REFERENCE"	
    
    arrHeader(0) = "VAT유형"						
    arrHeader(1) = "VAT유형명"						
    arrHeader(2) = "VAT율"
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
	Call vspdData_Change(frm1.vspdData.Col, frm1.vspdData.Row)		' 변경이 일어났다고 알려줌 

	lgBlnFlgChgValue = True
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
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	IsOpenPop = True 

	arrParam(0) = "Cost Center 팝업"			' 팝업 명칭 
	arrParam(1) = "B_COST_CENTER"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCostCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "B_COST_CENTER.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				" AND B_COST_CENTER.COST_TYPE ='M'" & _
				" AND B_COST_CENTER.DI_FG ='D'"			' Where Condition
	arrParam(5) = "Cost Center"					' TextBox 명칭 
	
    arrField(0) = "COST_CD"							' Field명(0)
    arrField(1) = "COST_NM"							' Field명(1)
    
    arrHeader(0) = "Cost Center"				' Header명(0)
    arrHeader(1) = "Cost Center 명"				' Header명(1)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCostCtr(arrRet)
	End If	
    
End Function

'------------------------------------------  SetWc()  --------------------------------------------------
'	Name : SetWc()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetWc(Byval arrRet)
	Dim lRow
	With frm1
		lRow = .vspdData.ActiveRow 
		ggoSpread.Source = .vspdData 
		.vspdData.Col = C_WcCd
		.vspdData.Row = .vspdData.ActiveRow 
		
		.vspdData.Text = arrRet(0)
		
		.vspdData.Col = C_HiddenInsideFlg
		.vspdData.Text = UCase(arrRet(2))
		
		If UCase(arrRet(2)) = "Y" then
			.vspdData.Col = C_InsideFlg 
			.vspdData.Text = "사내"
			
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
			.vspdData.Col = C_InsideFlg
			.vspdData.Text = "외주"
			
			.vspdData.ReDraw = False
			
			Call SetFieldProp(lRow,"N")
			
			.vspdData.ReDraw = True
		End if			
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' 변경이 일어났다고 알려줌 
	
	End With
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Of Measure Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(Byval arrRet)
	With frm1
		.vspdData.Col = C_RunTimeUnit
		.vspdData.Text = UCase(arrRet(0))
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' 변경이 일어났다고 알려줌 
	End With
End Function

'------------------------------------------  SetBizPartner()  --------------------------------------------------
'	Name : SetBizPartner()
'	Description : RoutingNo Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizPartner(Byval arrRet)
	With frm1.vspdData
		Call .SetText(C_BpCd, .ActiveRow, UCase(arrRet(0))) 
		Call .SetText(C_BpNm, .ActiveRow, UCase(arrRet(1))) 
		Call .SetText(C_CurCd, .ActiveRow, UCase(arrRet(3))) 
		Call .SetText(C_TaxType, .ActiveRow, UCase(arrRet(4))) 
		Call vspdData_Change(0, .Row)	' 변경이 일어났다고 알려줌 
	End With
End Function

'------------------------------------------  SetCurrency()  --------------------------------------------------
'	Name : SetCurrency()
'	Description : RoutingNo Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrency(Byval arrRet)
	With frm1
		.vspdData.Col = C_CurCd
		.vspdData.Text = UCase(arrRet(0))
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' 변경이 일어났다고 알려줌 
	
	End With
End Function

'------------------------------------------  SetRoutingNo()  --------------------------------------------------
'	Name : SetRoutingNo()
'	Description : RoutingNo Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRoutingNo(Byval arrRet)
	frm1.txtRoutingNo.value = arrRet(0)
	frm1.txtRoutingNm.value = arrRet(1)
End Function

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet,ByVal iPos)
	With frm1
		If iPos = 0 Then
			.txtItemCd.value = arrRet(0)
			.txtItemNm.value = arrRet(1)
		Else
			.txtItemCd1.value = arrRet(0)
			.txtItemNm1.value = arrRet(1)
			Call LookUpItemBasicUnit()
			lgBlnFlgChgValue = True
		End If

	End With
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
	Call txtPlantCd_OnChange()
End Function

'------------------------------------------  SetRoutCopy()  --------------------------------------------------
'	Name : SetRoutCopy()
'	Description : Routing Copy Reference Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRoutCopy(byval arrRet)
		
	Dim strVal
    
    LayerShowHide(1)
		
    strVal = BIZ_PGM_COPY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&RoutNo=" & arrRet(0)				'☆: 조회 조건 데이타 
    strVal = strVal & "&lgCurDt=" & StartDate
    strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	 
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

End Function

'------------------------------------------  SetRoutCopyOk()  --------------------------------------------------
'	Name : SetRoutCopy()
'	Description : Routing Copy Reference Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRoutCopyOk(ByVal LngMaxRow)
	Dim iRow
	Dim intindex
	
	frm1.vspdData.ReDraw = False

	Call InitData(LngMaxRow)							'공정작업명 구함.

	With frm1.vspdData 
				
		For iRow = LngMaxRow To .MaxRows
			.Row = iRow
			.Col = 0
			.Text = ggoSpread.InsertFlag 
			
			ggoSpread.SpreadUnLock C_OprNo,iRow, C_OprNo,iRow
			ggoSpread.SpreadUnLock C_ValidFromDt, iRow, C_ValidFromDt, iRow						
			
			.Col = C_HiddenInsideFlg
			.Row = iRow 
			
			If UCase(.Text) = "N" Then
				Call SetSpreadColor(iRow, iRow, "N")
				Call SetFieldProp(iRow, "N")
			Else
				Call SetSpreadColor(iRow, iRow, "Y")
			End If
			
			.Row = iRow
			.Col = C_MilestoneFlg
			.Text = "N"
			.Col = C_InspFlg
			.Text = "N"
			.Col = C_RunTimeUnit
			.Text = lgItemBaseUnit
		Next
		
		Call ProtectMilestone(0)
		
	End With
	
	Call SetActiveCell(frm1.vspdData,1,LngMaxRow,"M","X","X")
	Set gActiveElement = document.activeElement	
	
	frm1.vspdData.ReDraw = True
	
End Function

'------------------------------------------  SetCostCtr()  -----------------------------------------------
'	Name : SetCostCtr()
'	Description : Cost Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCostCtr(byval arrRet)
	frm1.txtCostCd.value = arrRet(0)
	frm1.txtCostNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function JumpAllocComp()
    Dim IntRetCd, strVal
    
    '-----------------------
    'Precheck area
    '-----------------------
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
		
    If frm1.vspdData.ActiveRow <= 0 Then 
		Call DisplayMsgBox("181216", "X", "X", "X")
		Exit Function
	End If
		
	WriteCookie "txtPlantCd", UCase(Trim(frm1.txtPlantCd.value))
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	WriteCookie "txtItemCd", UCase(Trim(frm1.txtItemCd1.value))
	WriteCookie "txtItemNm", frm1.txtItemNm1.value 
	WriteCookie "txtRoutingNo", UCase(Trim(frm1.txtRoutingNo1.value))
	WriteCookie "txtRoutingNm", frm1.txtRoutingNm1.value  
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_OprNo
	
	WriteCookie "txtOprNo", UCase(Trim(frm1.vspdData.Text))

	PgmJump(BIZ_PGM_JUMPCOMPALLOC_ID)	
	
End Function

Sub LookUpBp(ByVal pBpCd)	'2003-08-29

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strSelect, strFrom, strWhere
	Dim arrRows0, arrRows1, arrRows2, arrRows3

	If Trim(pBpCd) = "" Then Exit Sub

	'----------------------------------------------------------------------------------
	strSelect	= " BP_CD, " & _
				" BP_NM, " & _
				" CURRENCY, " & _
				" VAT_TYPE "
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

Sub LookUpWc(ByVal Str, ByVal Row)
	Dim strVal
	
	LayerShowHide(1)
		
    strVal = BIZ_PGM_LOOKUPWC_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtWcCd=" & Trim(Str)					
	strVal = strVal & "&Row=" & Row								
	    
	Call RunMyBizASP(MyBizASP, strVal)							
	
End Sub

Function LookUpWcOk(ByVal WcNm,ByVal InsideFlg, ByVal Row)
	With frm1.vspdData
		.ReDraw = False
		
		.Row = CLng(Row)
		.Col = C_HiddenInsideFlg
		.Text = InsideFlg
		
		If UCase(InsideFlg) = "Y" Then
			
			.Col = C_InsideFlg
			.Text = "사내"

			.Col = C_BpCd
			.Text = ""
			.Col = C_BpNm
			.Text = ""			
			.Col = C_CurCd
			.Text = ""
			.Col = C_SubconPrc
			.Text = ""
			.Col = C_TaxType
			.Text = ""
			
			Call SetFieldProp(Row, "Y")
			
		Else
			.Col = C_InsideFlg
			.Text = "외주"
			
			Call SetFieldProp(Row, "N")
		End If				
		
		.ReDraw = True
		
	End With
	IsOpenPop = False
	
End Function

Function LookUpWcNotOk(ByVal Row)
	Call SheetFocus(Row, C_WcCd)
	IsOpenPop = False
End Function

Sub ProtectCostCd()
	If UCase(Trim(Frm1.hOprCostFlag.value)) = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtCostCd, "N")  
	Else
		Frm1.txtCostCd.value = ""
		Frm1.txtCostNm.value = ""
		Call ggoOper.SetReqAttr(frm1.txtCostCd, "Q")  
	End If
End Sub

Sub ProtectMilestone(ByVal pvFlag)
	Dim iIntCnt
	Dim iStrFlag
	Dim iStrMilestoneFlg
	Dim iStrInspFlg

	ggoSpread.SpreadUnLock 	C_MilestoneFlg, 1, C_MilestoneFlg, frm1.vspdData.MaxRows
	ggoSpread.SpreadUnLock 	C_InspFlg, 1, C_InspFlg, frm1.vspdData.MaxRows
	ggoSpread.SSSetRequired 	C_MilestoneFlg, 1, frm1.vspdData.MaxRows
	ggoSpread.SSSetRequired 	C_InspFlg, 1, frm1.vspdData.MaxRows
	ggoSpread.SSSetRequired C_RunTimeUnit, 	-1, -1
	
	For iIntCnt = frm1.vspdData.MaxRows To 1 Step -1
		Call frm1.vspdData.GetText(0, iIntCnt, iStrFlag)
			
		Select Case iStrFlag
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
				Call frm1.vspdData.SetText(C_MilestoneFlg, iIntCnt, "Y")
				Call frm1.vspdData.SetText(C_InspFlg, iIntCnt, "N")
				ggoSpread.SSSetProtected C_MilestoneFlg, iIntCnt, iIntCnt
				ggoSpread.SSSetProtected C_InspFlg, iIntCnt, iIntCnt
				Exit For

			Case ""
				Call frm1.vspdData.GetText(C_MilestoneFlg, iIntCnt, iStrMilestoneFlg)
				Call frm1.vspdData.GetText(C_InspFlg, iIntCnt, iStrInspFlg)

				If iStrMilestoneFlg = "N" Or C_InspFlg = "Y" Then
					Call frm1.vspdData.SetText(0, iIntCnt, ggoSpread.UpdateFlag)
				End If
				Call frm1.vspdData.SetText(C_MilestoneFlg, iIntCnt, "Y")
				Call frm1.vspdData.SetText(C_InspFlg, iIntCnt, "N")

				ggoSpread.SSSetProtected C_MilestoneFlg, iIntCnt, iIntCnt
				ggoSpread.SSSetProtected C_InspFlg, iIntCnt, iIntCnt
				Exit For
		End Select
	Next
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtALTRTVALUE_Change
'   Event Desc : 
'=======================================================================================================
Sub txtALTRTVALUE_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
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
'   Event Desc : 달력을 호출한다.
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

'======================================================================================================
'   Event Name : txtPlantCd_OnChange
'   Event Desc : 검색창 공장코드가 변경될 경우 
'=======================================================================================================
Function txtPlantCd_OnChange()
    Dim IntRetCd

    If  frm1.txtPlantCd.value = "" Then
        frm1.txtPlantCd.Value = ""
        frm1.txtPlantNm.Value = ""
        frm1.hOprCostFlag.value = ""
    Else
		
        IntRetCD =  CommonQueryRs(" a.plant_nm, b.opr_cost_flag "," b_plant a (nolock), p_plant_configuration b (nolock) ", _
							" a.plant_cd = b.plant_cd and a.plant_cd = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & "" , _
							lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False   Then
			frm1.txtPlantNm.Value=""
			frm1.hOprCostFlag.value = ""
        Else
            frm1.txtPlantNm.Value= Trim(Replace(lgF0,Chr(11),""))
            frm1.hOprCostFlag.Value= Trim(Replace(lgF1,Chr(11),""))
        End If
		
     End If
     
     Call ProtectCostCd()
     
End Function

'=======================================================================================================
'   Event Name : txtItemCd1_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtItemCd1_OnChange()	
	lgBlnFlgChgValue = True	
	Call LookUpItemBasicUnit()
End Sub	

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	
	Select Case Col 
		Case C_WcCd
			If Trim(GetSpreadText(frm1.vspdData,Col,Row,"X","X")) <> "" Then
				Call LookUpWc(Trim(GetSpreadText(frm1.vspdData,Col,Row,"X","X")), Row)
			End If
			IsOpenPop = True
		Case C_OprNo
			If Trim(GetSpreadText(frm1.vspdData,Col,Row,"X","X")) <> "" Then
				If CheckValidOprNo(Trim(GetSpreadText(frm1.vspdData,Col,Row,"X","X")), Row) = False Then
					Call frm1.vspdData.SetText(Col,Row,"")
					Exit Sub
				End If
			End If
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

End Sub


'==========================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'==========================================================================================

Sub vspdData_EditChange(ByVal Col , ByVal Row )
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )	
    'Call SetPopupMenuItemInf("1101110111")
    	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001110111")
	Else 	
		If frm1.vspdData.MaxRows = 0 Then 
			Call SetPopupMenuItemInf("1001110111")
		Else
			Call SetPopupMenuItemInf("1101110111") 
		End if			
	End If	
	
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Or Col < 1 Then                                                    'If there is no data.
       Exit Sub
   	End If

End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_SubconPrc
            Call EditModeCheck(frm1.vspdData, Row, C_CurCd, C_SubconPrc, "C" ,"I", Mode, "X", "X")        
    End Select
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
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
		
    ElseIf Row >0 And Col = C_UnitPopup Then
		.Col = C_RunTimeUnit
        .Row = Row
        
        Call OpenUnit(.Text)
        Call SetActiveCell(frm1.vspdData,C_RunTimeUnit,Row,"M","X","X")
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
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspddata_KeyPress(index , KeyAscii)
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
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
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
    if frm1.vspdData.MaxRows < NewTop +  VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
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

'================================================ txtItem_onChange() ===============================
'   Event Name : txtItem_onChange()
'   Event Desc : 
'==========================================================================================
Sub txtItemCd1_OnChange()
	'Call LookUpItemByPlant
	'IsOpenPop = True
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
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
	
	If frm1.txtRoutingNo.value = "" Then
		frm1.txtRoutingNm.value = ""
	End If
		
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
    
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
    End If     									'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
   
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    Dim slPlantCd
    Dim slPlantNm
    
    FncNew = False                                                          '⊙: Processing is NG
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    slPlantCd = frm1.txtPlantCd.value
    slPlantNm = frm1.txtPlantNm.value

    
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    
    frm1.txtPlantCd.value = slPlantCd
    frm1.txtPlantNm.value = slplantNm
    
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call txtPlantCd_OnChange()
    
    Call SetToolbar("11101101001011")
    
    frm1.txtItemCd1.focus 
    Set gActiveElement = document.activeElement 
    
    FncNew = True                                                           '⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '⊙: "Will you destory previous data"	
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
    If DbDelete = False Then   
		Exit Function           
    End If     						'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK

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
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
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
    
    On Error Resume Next 
    With frm1.vspdData
		If .MaxRows < 1 Then Exit Function
    
		.Focus

		.EditMode = True
	
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
    
		ggoSpread.CopyRow
		Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData, .ActiveRow, .ActiveRow, C_CurCd, C_SubconPrc, "C", "I", "X", "X")
		.ReDraw = True
		
		.ReDraw = False
		.Col = C_HiddenInsideFlg
		.Row = .ActiveRow
	
		If UCase(.Text) = "N" Then
			Call SetSpreadColor(.ActiveRow, .ActiveRow, "N") 
		Else
			Call SetSpreadColor(.ActiveRow, .ActiveRow, "Y") 
		End If
    
		'------------------------------------------------------
		' Default Value Setting
		'------------------------------------------------------
		.Row = .ActiveRow
		.Col = C_OprNo
		.Text = ""
    
		.Col = C_ValidFromDt
		.Text = StartDate
    
		.Col = C_ValidToDt
		.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")

		Call ProtectMilestone(0)
		    
		.ReDraw = True
	End With
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

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_CurCd,C_SubconPrc, "C" ,"I","X","X")
	
	frm1.vspdData.Col = C_HiddenInsideFlg
	
	If UCase(Trim(frm1.vspdData.Text)) = "N" Then
		Call SetFieldProp(frm1.vspdData.Row, "N")
	Else
		Call SetFieldProp(frm1.vspdData.Row, "Y")
	End IF
    
	frm1.vspdData.Redraw = False
	Call InitData(1)
	Call ProtectMilestone(0)
	frm1.vspdData.Redraw = True
	
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
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

    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If
	
    With frm1
 
        .vspdData.ReDraw = False
		.vspdData.Focus
    
		ggoSpread.Source = .vspdData

        If frm1.vspdData.selBlockRow = -1 Then
            ggoSpread.InsertRow 0, iIntReqRows
        Else
            ggoSpread.InsertRow , iIntReqRows
        End If
	
	    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + iIntReqRows - 1,C_CurCd,C_SubconPrc, "C" ,"I","X","X")
	
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1, "Y")
		For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + iIntReqRows - 1
			.vspdData.Row = iIntCnt
			.vspdData.Col = C_MilestoneFlg
			.vspdData.Text = "N"    
			.vspdData.Col = C_InspFlg
			.vspdData.Text = "N"    
			.vspdData.Col = C_ValidFromDt
			.vspdData.Text = StartDate
			.vspdData.Col = C_ValidToDt
			.vspdData.Text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
			.vspdData.Col = C_RunTimeUnit
			.vspdData.Text = lgItemBaseUnit
		Next

		Call ProtectMilestone(0)

        .vspdData.ReDraw = True

	End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 
    Dim iIntCnt
    Dim iChrFlag

    '----------------------
    ' 데이터가 없는 경우 
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
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                                         '☜:화면 유형, Tab 유무 
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

    DbDelete = False														'⊙: Processing is NG
    
    LayerShowHide(1)
		
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtRoutingNo=" & Trim(frm1.txtRoutingNo.value)				'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    DbQuery = False                                                         '⊙: Processing is NG
    
    LayerShowHide(1)
		
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001								'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtRoutingNo=" & Trim(frm1.hRoutingNo.value)				'☆: 조회 조건 데 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtRoutingNo=" & Trim(frm1.txtRoutingNo.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=0"
    End If

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)													'☆: 조회 성공후 실행로직 
	Dim i
    '-----------------------
    'Reset variables area
    '-----------------------
    frm1.vspdData.redraw = False
    
    Call InitData(LngMaxRow)
    	
    For i= LngMaxRow To frm1.vspdData.MaxRows
		frm1.vspdData.Col = C_HiddenInsideFlg
		frm1.vspdData.Row = i
	
		If UCase(Trim(frm1.vspdData.Text)) = "N" Then
			Call SetFieldProp(i, "N")
		Else
			Call SetFieldProp(i, "Y")
		End IF
		
    Next
    
    Call LookUpItemBasicUnit()
	Call ProtectMilestone(1)
		
	frm1.vspdData.redraw = True
	
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field
	Call ProtectCostCd()
	
    If frm1.vspdData.MaxRows = 0 Then 
		Call SetToolbar("11111101001111")
	Else
		Call SetToolbar("11111111001111")
	End if	
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
	 
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows 
    Dim strVal
	Dim strDel
	Dim strOprNo, strInsideFlg, strMilestoneFlg, strInspFlg, strValidFromDt, strValidToDt, strQueueTime, strSetupTime, strWaitTime, strFixRunTime
	Dim strRunTime, strMoveTime, strOverLapOpr, strSubconPrc
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
	
	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size

	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size
		
    DbSave = False                                                          '⊙: Processing is NG
    
    If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function      

    LayerShowHide(1)
		
    'On Error Resume Next                                                   '☜: Protect system from crashing

    With frm1
		.txtMode.Value = parent.UID_M0002											'☜: 저장 상태 
		.txtFlgMode.Value = lgIntFlgMode									'☜: 신규입력/수정 상태 
		
		'--------------------------------
		' 신규입력이면 무조건 BOM Type =1
		'--------------------------------
		If lgIntFlgMode = parent.OPMD_CMODE Then
			.txtBomNo.value = "1"
		End If
		
	End With
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0
	
	With frm1.vspdData
	    
    For IntRows = 1 To .MaxRows
    
		.Row = IntRows
		.Col = 0

		Select Case .Text
	    
		    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
				
				strVal = ""
							
				If .Text = ggoSpread.InsertFlag Then
					strVal = strVal & "C" & iColSep & IntRows & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
				Else
					strVal = strVal & "U" & iColSep	& IntRows & iColSep				'⊙: U=Update
				End If
				
		        .Col = C_OprNo								'2
		        strVal = strVal & Trim(.Text) & iColSep
		        strOprNo = Trim(.Text)
		            
		        .Col = C_HiddenRoutOrder					'3
				strVal = strVal & Trim(.Text) & iColSep
		        
		        .Col = C_WCCd								'4
		        strVal = strVal & Trim(.Text) & iColSep

		        .Col = C_HiddenInsideFlg					'5
		        strVal = strVal & Trim(.Text) & iColSep
		        strInsideFlg = Trim(.Text)
		        
		        .Col = C_MilestoneFlg						'6
		        strVal = strVal & Trim(.Text) & iColSep
		        strMilestoneFlg = Trim(.Text)

    			.Col = C_RunTimeQty							'7
		        strVal = strVal & UNIConvNum(Trim(.Text), 0) & iColSep
		        
		        .Col = C_RunTimeUnit						'8
		        strVal = strVal & UCase(Trim(.Text)) & iColSep
		        
				.Col = C_InspFlg							'9
		        strInspFlg = Trim(.Text)
				If Trim(strMilestoneFlg) = "N" And Trim(strInspFlg) = "Y" Then
					Call DisplayMsgBox("181217", "X", "X", "X")
					Call SheetFocus(IntRows, C_InspFlg)
					Exit Function
				Else
					strVal = strVal & Trim(.Text) & iColSep
				End If
		        		        
		        .Col = C_ValidFromDt						'10
		        strValidFromDt = Trim(.Text)
				If Len(Trim(strValidFromDt)) Then
					If UNIConvDate(strValidFromDt) = "" Then	 
						Call DisplayMsgBox("122116", "X", "X", "X")
						Call SheetFocus(IntRows, C_ValidToDt)
						Exit Function
					Else
						strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					End If
				End If
				
				.Col = C_ValidToDt							'11
		        strValidToDt = Trim(.Text)
				If Len(Trim(strValidToDt)) Then
					If UNIConvDate(strValidToDt) = "" Then	 
						Call DisplayMsgBox("122116", "X", "X", "X")
						Call SheetFocus(IntRows, C_ValidToDt)
						Exit Function
					Else
						strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					End If
				End If

				If UNIConvDate(strValidFromDt) > UNIConvDate(strValidToDt) Then
					Call DisplayMsgBox("972002", "X", "종료일", "시작일")
					Call SheetFocus(IntRows, C_ValidToDt)
					Exit Function
				End If					
		        
		        .Col = C_JobCd								'12
		        strVal = strVal & Trim(.Text) & iColSep

		        .Col = C_MfgLt								'13
		        strVal = strVal & UNIConvNum(Trim(.Text), 0) & iColSep

		        .Col = C_QueueTime							'14
		        strQueueTime = Trim(.Text)
				If ConvToSec(strQueueTime) = -999999	Then
					Call DisplayMsgBox("970029", "X", "Queue Time", "X")
					Call SheetFocus(IntRows, C_QueueTime)
					Exit Function	
				Else
					strVal = strVal & ConvToSec(strQueueTime) & iColSep
				End If

		        .Col = C_SetupTime							'15
		        strSetupTime = Trim(.Text)
				If ConvToSec(strSetupTime) = -999999	Then
					Call DisplayMsgBox("970029", "X", "설치시간", "X")
					Call SheetFocus(IntRows, C_SetupTime)
					Exit Function	
				Else
					strVal = strVal & ConvToSec(strSetupTime) & iColSep
				End If

		        .Col = C_WaitTime							'16
		        strWaitTime = Trim(.Text)
				If ConvToSec(strWaitTime) = -999999	Then
					Call DisplayMsgBox("970029", "X", "대기시간", "X")
					Call SheetFocus(IntRows, C_WaitTime)
					Exit Function	
				Else
					strVal = strVal & ConvToSec(strWaitTime) & iColSep
				End If

		        .Col = C_FixRunTime							'16
		        strFixRunTime = Trim(.Text)
				If ConvToSec(strFixRunTime) = -999999	Then
					Call DisplayMsgBox("970029", "X", "고정가동시간", "X")
					Call SheetFocus(IntRows, C_FixRunTime)
					Exit Function	
				Else
					strVal = strVal & ConvToSec(strFixRunTime) & iColSep
				End If
    
		        .Col = C_RunTime							'17
		        strRunTime = Trim(.Text)
				If ConvToSec(strRunTime) = -999999	Then
					Call DisplayMsgBox("970029", "X", "변동가동시간", "X")
					Call SheetFocus(IntRows, C_RunTime)
					Exit Function	
				Else
					strVal = strVal & ConvToSec(strRunTime) & iColSep
				End If
    			
		        .Col = C_MoveTime							'18
		        strMoveTime = Trim(.Text)
				If ConvToSec(strMoveTime) = -999999	Then
					Call DisplayMsgBox("970029", "X", "이동시간", "X")
					Call SheetFocus(IntRows, C_MoveTime)
					Exit Function	
				Else
					strVal = strVal & ConvToSec(strMoveTime) & iColSep
				End If
		        
		        .Col = C_OverLapOpr							'19
		        strOverLapOpr = Trim(.Text)
				If Trim(strOverLapOpr) <> "" Then
					If CheckValidOverlapOpr(Trim(strOverLapOpr), IntRows) = False Then
						Call SheetFocus(IntRows, C_OverLapOpr)
						Exit Function
					End If	
				End If
				
				strVal = strVal & Trim(.Text) & iColSep
		        
		        .Col = C_OverLapLt							'20
		        strVal = strVal & UNIConvNum(Trim(.Text), 0) & iColSep
		        
		        .Col = C_BpCd								'21
		        strVal = strVal & Trim(.Text) & iColSep
		        
		        .Col = C_SubconPrc							'22
		        strSubconPrc = Trim(.Text)
				If strInsideFlg = "N" And UNIConvNum(strSubconPrc,0) = 0 Then
					Call DisplayMsgBox("970022", "X" , "공정외주단가", "0")
					Call SheetFocus(IntRows, C_SubconPrc)
					Exit Function
				End If
		        strVal = strVal & UNIConvNum(Trim(.Text), 0) & iColSep
		        
		        .Col = C_CurCd								'23
		        strVal = strVal & Trim(.Text) & iColSep
		        
		        .Col = C_TaxType							'24
		        strVal = strVal & UCase(Trim(.Text)) & iRowSep		        
		        
		    Case ggoSpread.DeleteFlag
				
				strDel = ""
				
				strDel = strDel & "D" & iColSep	& IntRows & iColSep				'⊙: D=Delete
				
				.Col = C_OprNo	'2
				If Trim(.Text) <> "" Then
					If CheckOverlapOprExist(Trim(.Text), IntRows) = True Then
						Call SheetFocus(IntRows, C_OprNo)
						Exit Function
					End If
				End If
				
		        strDel = strDel & Trim(.Text) & iRowSep
		        
		End Select
		
		.Col = 0
		Select Case .Text
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			    
		         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
			 
		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
			       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
			         
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			         
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
			          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
			       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
			         
		         iTmpDBuffer(iTmpDBufferCount) =  strDel         
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
		End Select
		
    Next

	End With
	
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 
    DbSave = True                                                           '⊙: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	
	frm1.txtItemCd.value = frm1.txtItemCd1.value
	frm1.txtRoutingNo.value = frm1.txtRoutingNo1.value 
		
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()
	IsOpenPop = False
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

Function CheckValidOprNo(ByVal pvOprNo, ByVal pvRow)
	Dim iIntCnt, iStrPrevOprNo, iStrNextOprNo
	Dim iStrValue
	
	CheckValidOprNo = True
	iStrPrevOprNo = "" : iStrNextOprNo = "z"

	For iIntCnt = pvRow - 1 To 1 Step -1
		Call frm1.vspdData.GetText(0, iIntCnt, iStrValue)
		If iStrValue <> ggoSpread.DeleteFlag Then
			Call frm1.vspdData.GetText(C_OprNo, iIntCnt, iStrPrevOprNo)
			If iStrPrevOprNo <> "" Then
				Exit For
			End If
		End If
	Next

	For iIntCnt = pvRow + 1 To frm1.vspdData.MaxRows Step 1
		Call frm1.vspdData.GetText(0, iIntCnt, iStrValue)
		If iStrValue <> ggoSpread.DeleteFlag Then
			Call frm1.vspdData.GetText(C_OprNo, iIntCnt, iStrNextOprNo)
			If iStrNextOprNo <> "" Then
				Exit For
			Else
				iStrNextOprNo = "z"
			End If
		End If
	Next
		
	If pvOprNo >= iStrNextOprNo Or pvOprNo <= iStrPrevOprNo Then
		If iStrPrevOprNo = "" Then
			Call DisplayMsgBox("181220", "X", iStrNextOprNo, "X")
		ElseIf iStrNextOprNo = "z" Then
			Call DisplayMsgBox("181219", "X", iStrPrevOprNo, "X")
		Else
			Call DisplayMsgBox("181218", "X", iStrPrevOprNo, iStrNextOprNo)
		End If
		CheckValidOprNo = False
	End If
	
End Function

Function CheckValidOverlapOpr(ByVal pvOverlapOpr, ByVal pvRow)
	Dim iIntCnt, iStrValue

	CheckValidOverlapOpr = False
	
	For iIntCnt = pvRow - 1 To 1 Step -1
		Call frm1.vspdData.GetText(0, iIntCnt, iStrValue)

		If iStrValue <> ggoSpread.DeleteFlag Then
			Call frm1.vspdData.GetText(C_OprNo, iIntCnt, iStrValue)
			If iStrValue <> pvOverlapOpr Then
				CheckValidOverlapOpr = False
				Call DisplayMsgBox("181318", "X", "X", "X")        'Overlap공정이 일치하지 않음 
			Else
				CheckValidOverlapOpr = True
			End If

			Exit Function
		End If
	Next

	Call DisplayMsgBox("181316", "X", "X", "X")                    '초공정이 Overlap공정을 가짐 
End Function

Function CheckOverlapOprExist(ByVal pvOverlapOpr, ByVal pvRow)
	Dim iIntCnt, iStrValue

	CheckOverlapOprExist = False

	For iIntCnt = pvRow + 1 To frm1.vspdData.MaxRows
		Call frm1.vspdData.GetText(0, iIntCnt, iStrValue)
		If iStrValue <> ggoSpread.DeleteFlag Then
			Call frm1.vspdData.GetText(C_OverlapOpr, iIntCnt, iStrValue)
			If iStrValue = pvOverlapOpr Then
				CheckOverlapOprExist = True
				Call DisplayMsgBox("181319", "X", "X", "X")        'Overlap공정이 존재함 
				Exit Function
			End If
		End If
	Next
End Function

'==============================================================================
' Function : LookUpItemBasicUnit()
' Description : 라우팅 정보의 품목 기준단위를 가져옴,
'==============================================================================
Function LookUpItemBasicUnit()

	If Trim(frm1.txtItemCd1.value) <> "" Then
		If 	CommonQueryRs(" BASIC_UNIT "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd1.value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then		
			lgItemBaseUnit = ""
		Else
			lgF0 = Split(lgF0, Chr(11))
			lgItemBaseUnit = Trim(lgF0(0))
		End If
	Else
		lgItemBaseUnit = ""
	End If	

End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
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
