Const BIZ_PGM_INIT_QRY_ID	= "p1713mb11.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_QRY_ID		= "p1713mb12.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID		= "p1713mb13.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID		= "p1713mb14.asp"			'☆: 비지니스 로직 ASP명 

Dim C_Level
Dim C_Seq
Dim C_ChildItemCd
Dim C_ChildItemPopUp
Dim C_ChildItemNm
Dim C_Spec
Dim C_ChildItemUnit
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_ProcType
Dim C_ProcTypeNm
Dim C_BomType
Dim C_BomTypePopup
Dim C_ChildItemBaseQty
Dim C_ChildBasicUnit
Dim C_ChildBasicUnitPopup
Dim C_PrntItemBaseQty
Dim C_PrntBasicUnit
Dim C_PrntBasicUnitPopup
Dim C_SafetyLT
Dim C_LossRate
Dim C_SupplyFlg
Dim C_SupplyFlgNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_ECNNo
Dim C_ECNNoPopup
Dim C_ECNDesc
Dim C_ReasonCd
Dim C_ReasonCdPopup
Dim C_ReasonNm
Dim C_DrawingPath
Dim C_Remark
Dim C_HdrItemCd
Dim C_HdrBomNo
Dim C_HdrProcType
Dim C_ItemValidFromDt
Dim C_ItemValidToDt
Dim C_ItemAcctGrp
Dim C_ReqTransNo
Dim C_ReqTransDt
Dim C_TransStatus
Dim C_TransDt
Dim C_Row
Dim C_Row1


Dim isClicked
Dim iCol
Dim iRow
Dim IsOpenPop
Dim lgStrBOMHisFlg
Dim iStrFree

Dim lgStrPrevKeyIndex1


'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

	C_Level					= 1
	C_Seq					= 2
	C_ChildItemCd			= 3
	C_ChildItemPopUp		= 4
	C_ChildItemNm			= 5
	C_Spec					= 6
	C_ChildItemUnit			= 7
	C_ItemAcct				= 8
	C_ItemAcctNm			= 9
	C_ProcType				= 10
	C_ProcTypeNm			= 11
	C_BomType				= 12
	C_BomTypePopup			= 13
	C_ChildItemBaseQty		= 14
	C_ChildBasicUnit		= 15
	C_ChildBasicUnitPopup	= 16
	C_PrntItemBaseQty		= 17
	C_PrntBasicUnit			= 18
	C_PrntBasicUnitPopup	= 19
	C_SafetyLT				= 20
	C_LossRate				= 21
	C_SupplyFlg				= 22
	C_SupplyFlgNm			= 23
	C_ValidFromDt			= 24
	C_ValidToDt				= 25
	C_ECNNo					= 26
	C_ECNNoPopup			= 27
	C_ECNDesc				= 28
	C_ReasonCd				= 29
	C_ReasonCdPopup			= 30
	C_ReasonNm				= 31
	C_DrawingPath			= 32
	C_Remark				= 33
	C_HdrItemCd				= 34
	C_HdrBomNo				= 35
	C_HdrProcType			= 36
	C_ItemValidFromDt		= 37
	C_ItemValidToDt			= 38
	C_ItemAcctGrp			= 39
	
	'vspdData1
	C_Row1					= 40
	
	'vspdData
	C_ReqTransNo			= 40
	C_ReqTransDt			= 41
	C_TransStatus			= 42
	C_TransDt				= 43
	C_Row					= 44	

		
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKeyIndex = ""
    lgStrPrevKeyIndex1 = ""    
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                               '⊙: initializes sort direction

End Sub

Sub SetDefaultVal()
	frm1.txtBaseDt.value = StartDate
	frm1.hStartDate.value = StartDate
	lgStrBOMHisFlg = "N"
	
	frm1.hStatus.value = "N"
	frm1.hRequestingFlg.value = "N"		'이관의뢰버턴을 클릭하여 이관의뢰작업중인경우에 'Y'
	frm1.hBomType.value = "1"
	
	'기본적으로 쿼리는 설계BOM과 제조BOM을 동시에 조회한다.
	frm1.txtQueryType.value = "A"
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	'============================================================================================
	'☜: Spreadsheet vspdData (하단부)
	'============================================================================================
	
		ggoSpread.Source = frm1.vspdData
	
		With frm1.vspdData
	
			ggoSpread.Spreadinit "V20050125",, parent.gAllowDragDropSpread
	
			.ReDraw = False
	
			.MaxCols = C_Row
			.MaxRows = 0
	
			Call GetSpreadColumnPos("A")
	
			ggoSpread.SSSetEdit 	C_Level, 				"레벨", 8
			ggoSpread.SSSetFloat	C_Seq,					"순서", 6, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
			ggoSpread.SSSetEdit		C_ChildItemCd,			"자품목", 20,,, 18, 2
			ggoSpread.SSSetButton	C_ChildItemPopUp
			ggoSpread.SSSetEdit 	C_ChildItemNm, 			"자품목명", 30
			ggoSpread.SSSetEdit 	C_Spec,	 				"규격", 30
			ggoSpread.SSSetEdit		C_ChildItemUnit,		"단위", 6,,, 3, 2
			ggoSpread.SSSetEdit		C_ItemAcct,				"품목계정", 10
			ggoSpread.SSSetEdit		C_ItemAcctNm,			"품목계정", 10
			ggoSpread.SSSetEdit 	C_ProcType, 			"조달구분", 10
			ggoSpread.SSSetEdit 	C_ProcTypeNm, 			"조달구분", 12
			ggoSpread.SSSetEdit		C_BomType,				"BOM Type", 10,,, 3, 2
	 		ggoSpread.SSSetButton	C_BomTypePopup
			ggoSpread.SSSetFloat	C_ChildItemBaseQty,		"자품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
			ggoSpread.SSSetEdit 	C_ChildBasicUnit,		"단위"			, 6,,, 3, 2
			ggoSpread.SSSetButton	C_ChildBasicUnitPopup
			ggoSpread.SSSetFloat	C_PrntItemBaseQty,		"모품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
			ggoSpread.SSSetEdit		C_PrntBasicUnit,		"단위"			, 6,,, 3, 2
			ggoSpread.SSSetButton	C_PrntBasicUnitPopup
			ggoSpread.SSSetFloat 	C_SafetyLT, 			"안전L/T"	, 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
			ggoSpread.SSSetFloat	C_LossRate,				"Loss율"	, 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
			ggoSpread.SSSetCombo	C_SupplyFlg,			"유무상구분", 8
			ggoSpread.SSSetCombo	C_SupplyFlgNm,			"유무상구분", 10
			ggoSpread.SSSetDate		C_ValidFromDt,			"시작일"	, 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ValidToDt,			"종료일"	, 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_ECNNo,				"설계변경번호", 18,,, 18, 2
			ggoSpread.SSSetButton	C_ECNNoPopup
	 		ggoSpread.SSSetEdit		C_ECNDesc,				"설계변경내용", 30,,, 100
			ggoSpread.SSSetEdit		C_ReasonCd,				"설계변경근거", 10,,, 2, 2
			ggoSpread.SSSetButton	C_ReasonCdPopup
			ggoSpread.SSSetEdit		C_ReasonNm,				"설계변경근거명", 14
			ggoSpread.SSSetEdit		C_DrawingPath,			"도면경로", 30,,, 100
			ggoSpread.SSSetEdit 	C_Remark,	 			"비고"		, 30,,, 1000
			ggoSpread.SSSetEdit		C_HdrItemCd,			"Header품목", 5
			ggoSpread.SSSetEdit		C_HdrBomNo,				"header BOM No.", 5
			ggoSpread.SSSetEdit		C_HdrProcType,			"조달구분", 8
			ggoSpread.SSSetDate		C_ItemValidFromDt,		"품목시작일"	, 11, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_ItemValidToDt,		"품목종료일"	, 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_ItemAcctGrp,			"품목계정그룹", 10
			ggoSpread.SSSetEdit		C_ReqTransNo,			"이관의뢰번호", 12
			ggoSpread.SSSetDate		C_ReqTransDt,			"이관요청일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_TransStatus,			"이관상태", 10
			ggoSpread.SSSetDate		C_TransDt,				"이관일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_Row,					"순서", 5

			ggoSpread.SSSetSplit2(C_ChildItemPopUp)											'frozen 기능 추가 
		
			Call ggoSpread.MakePairsColumn(C_Level, C_ChildItemPopUp)
			Call ggoSpread.MakePairsColumn(C_BomType, C_BomTypePopup)
			Call ggoSpread.MakePairsColumn(C_ChildItemBaseQty, C_ChildBasicUnitPopup)
			Call ggoSpread.MakePairsColumn(C_PrntItemBaseQty, C_PrntBasicUnitPopup)
			Call ggoSpread.MakePairsColumn(C_ECNNo, C_ECNNoPopup)
			Call ggoSpread.MakePairsColumn(C_ReasonCd, C_ReasonCdPopup)
	
			Call ggoSpread.SSSetColHidden(C_ChildItemUnit, C_ChildItemUnit, True)
			Call ggoSpread.SSSetColHidden(C_ItemAcct, C_ItemAcct, True)
			Call ggoSpread.SSSetColHidden(C_ProcType, C_ProcType, True)
'			Call ggoSpread.SSSetColHidden(C_BomType, C_BomType, True)		'2005.01.24 E-BOM Modified 
			Call ggoSpread.SSSetColHidden(C_SupplyFlg, C_SupplyFlg, True)
			Call ggoSpread.SSSetColHidden(C_HdrItemCd, C_HdrItemCd, True)
			Call ggoSpread.SSSetColHidden(C_HdrBomNo, C_HdrBomNo, True)
			Call ggoSpread.SSSetColHidden(C_HdrProcType, C_HdrProcType, True)
			Call ggoSpread.SSSetColHidden(C_ItemValidFromDt, C_ItemValidFromDt, True)
			Call ggoSpread.SSSetColHidden(C_ItemValidToDt, C_ItemValidToDt, True)
			Call ggoSpread.SSSetColHidden(C_ItemAcctGrp, C_ItemAcctGrp, True)
			Call ggoSpread.SSSetColHidden(C_Row, C_Row, True)
			
			.ReDraw = True
			
	    End With
    
	'============================================================================================
	'☜: Spreadsheet vspdData1 (상단부)
	'============================================================================================
	ggoSpread.Source = frm1.vspdData1	

	With frm1.vspdData1

		ggoSpread.Spreadinit "V20050125",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_Row1
		.MaxRows = 0

		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetEdit 	C_Level, 				"레벨", 8
		ggoSpread.SSSetFloat	C_Seq,					"순서", 6, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetEdit		C_ChildItemCd,			"자품목", 20,,, 18, 2
		ggoSpread.SSSetButton	C_ChildItemPopUp
		ggoSpread.SSSetEdit 	C_ChildItemNm, 			"자품목명", 30
		ggoSpread.SSSetEdit 	C_Spec,	 				"규격", 30
		ggoSpread.SSSetEdit		C_ChildItemUnit,		"단위", 6,,, 3, 2
		ggoSpread.SSSetEdit		C_ItemAcct,				"품목계정", 10
		ggoSpread.SSSetEdit		C_ItemAcctNm,			"품목계정", 10
		ggoSpread.SSSetEdit 	C_ProcType, 			"조달구분", 10
		ggoSpread.SSSetEdit 	C_ProcTypeNm, 			"조달구분", 12
		ggoSpread.SSSetEdit		C_BomType,				"BOM Type", 10,,, 3, 2
 		ggoSpread.SSSetButton	C_BomTypePopup
		ggoSpread.SSSetFloat	C_ChildItemBaseQty,		"자품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
		ggoSpread.SSSetEdit 	C_ChildBasicUnit,		"단위"			, 6,,, 3, 2
		ggoSpread.SSSetButton	C_ChildBasicUnitPopup
		ggoSpread.SSSetFloat	C_PrntItemBaseQty,		"모품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,, "Z"
		ggoSpread.SSSetEdit		C_PrntBasicUnit,		"단위"			, 6,,, 3, 2
		ggoSpread.SSSetButton	C_PrntBasicUnitPopup
		ggoSpread.SSSetFloat 	C_SafetyLT, 			"안전L/T"	, 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetFloat	C_LossRate,				"Loss율"	, 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec, 1, FALSE, "Z" 
		ggoSpread.SSSetCombo	C_SupplyFlg,			"유무상구분", 8
		ggoSpread.SSSetCombo	C_SupplyFlgNm,			"유무상구분", 10
		ggoSpread.SSSetDate		C_ValidFromDt,			"시작일"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,			"종료일"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ECNNo,				"설계변경번호", 18,,, 18, 2
		ggoSpread.SSSetButton	C_ECNNoPopup
 		ggoSpread.SSSetEdit		C_ECNDesc,				"설계변경내용", 30,,, 100
		ggoSpread.SSSetEdit		C_ReasonCd,				"설계변경근거", 10,,, 2, 2
		ggoSpread.SSSetButton	C_ReasonCdPopup
		ggoSpread.SSSetEdit		C_ReasonNm,				"설계변경근거명", 14
		ggoSpread.SSSetEdit		C_DrawingPath,			"도면경로", 30,,, 100
		ggoSpread.SSSetEdit 	C_Remark,	 			"비고"		, 30,,, 1000
		ggoSpread.SSSetEdit		C_HdrItemCd,			"Header품목", 5
		ggoSpread.SSSetEdit		C_HdrBomNo,				"header BOM No.", 5
		ggoSpread.SSSetEdit		C_HdrProcType,			"조달구분", 8
		ggoSpread.SSSetDate		C_ItemValidFromDt,		"품목시작일"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ItemValidToDt,		"품목종료일"	, 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ItemAcctGrp,			"품목계정그룹", 10
		ggoSpread.SSSetEdit		C_Row1,					"순서", 5
								
		ggoSpread.SSSetSplit2(C_ChildItemPopUp)											'frozen 기능 추가 
	
		Call ggoSpread.MakePairsColumn(C_Level, C_ChildItemPopUp)
		Call ggoSpread.MakePairsColumn(C_BomType, C_BomTypePopup)
		Call ggoSpread.MakePairsColumn(C_ChildItemBaseQty, C_ChildBasicUnitPopup)
		Call ggoSpread.MakePairsColumn(C_PrntItemBaseQty, C_PrntBasicUnitPopup)
		Call ggoSpread.MakePairsColumn(C_ECNNo, C_ECNNoPopup)
		Call ggoSpread.MakePairsColumn(C_ReasonCd, C_ReasonCdPopup)

		Call ggoSpread.SSSetColHidden(C_ChildItemUnit, C_ChildItemUnit, True)
		Call ggoSpread.SSSetColHidden(C_ItemAcct, C_ItemAcct, True)
		Call ggoSpread.SSSetColHidden(C_ProcType, C_ProcType, True)
'		Call ggoSpread.SSSetColHidden(C_BomType, C_BomType, True)		'2005.01.24 E-BOM Modified 
		Call ggoSpread.SSSetColHidden(C_SupplyFlg, C_SupplyFlg, True)
		Call ggoSpread.SSSetColHidden(C_HdrItemCd, C_HdrItemCd, True)
		Call ggoSpread.SSSetColHidden(C_HdrBomNo, C_HdrBomNo, True)
		Call ggoSpread.SSSetColHidden(C_HdrProcType, C_HdrProcType, True)
		Call ggoSpread.SSSetColHidden(C_ItemValidFromDt, C_ItemValidFromDt, True)
		Call ggoSpread.SSSetColHidden(C_ItemValidToDt, C_ItemValidToDt, True)
		Call ggoSpread.SSSetColHidden(C_ItemAcctGrp, C_ItemAcctGrp, True)
		Call ggoSpread.SSSetColHidden(C_Row1, C_Row1, True)
		
		.ReDraw = True

    End With	'============================================================================================
    
	Call SetSpreadLock 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()

    With frm1
    	'하단부 Spread
    	ggoSpread.Source = frm1.vspdData
    	
		.vspdData.ReDraw = False

		ggoSpread.SSSetProtected	-1, -1
		
		If .hStatus.value = "N" Then
		'	ggoSpread.SpreadUnLock	C_ChildItemBaseQty,	-1, C_ChildItemBaseQty
		'	ggoSpread.SpreadUnLock	C_ChildBasicUnit,	-1, C_ChildBasicUnitPopup
		'	ggoSpread.SpreadUnLock	C_PrntItemBaseQty,	-1, C_PrntItemBaseQty
		'	ggoSpread.SpreadUnLock	C_PrntBasicUnit,	-1, C_PrntBasicUnitPopup
		'	ggoSpread.SpreadUnLock	C_SafetyLT,			-1, C_SafetyLT
		'	ggoSpread.SpreadUnLock	C_LossRate,			-1, C_LossRate
		'	ggoSpread.SpreadUnLock	C_ValidToDt,		-1, C_ValidToDt
		'	ggoSpread.SpreadUnLock	C_Remark,			-1, C_Remark
	
		'	ggoSpread.SSSetRequired C_ChildItemBaseQty, -1, -1
		'	ggoSpread.SSSetRequired C_ChildBasicUnit, 	-1, -1
		'	ggoSpread.SSSetRequired	C_PrntItemBaseQty,	-1, -1
		'	ggoSpread.SSSetRequired	C_PrntBasicUnit,	-1, -1
		'	ggoSpread.SSSetRequired C_ValidToDt, 		-1, -1
		End If
						
		.vspdData.ReDraw = True

    	'상단부 Spread
    	ggoSpread.Source = frm1.vspdData1
    	
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected	-1, -1
		.vspdData.ReDraw = True
				
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal Level, ByVal QueryStatus)
	
	Dim pvTmpRow
	
	ggoSpread.Source = frm1.vspdData

	frm1.vspdData.ReDraw = False

	If QueryStatus = 1 Then		'When Query is OK
	    If Level = 0 Then
			ggoSpread.SSSetProtected -1, pvStartRow, pvEndRow
		Else
	    	ggoSpread.SSSetProtected -1, pvStartRow, pvEndRow
			
			If frm1.hStatus.value = "N" Then
				If pvStartRow > 1 Then
					pvTmpRow = pvStartRow
				Else
					pvTmpRow = 2
				End If
				ggoSpread.SpreadUnLock	C_ChildItemBaseQty,	pvTmpRow, C_ChildItemBaseQty,		pvEndRow
				ggoSpread.SpreadUnLock	C_ChildBasicUnit,	pvTmpRow, C_ChildBasicUnitPopup,	pvEndRow
				ggoSpread.SpreadUnLock	C_PrntItemBaseQty,	pvTmpRow, C_PrntItemBaseQty,		pvEndRow
				ggoSpread.SpreadUnLock	C_PrntBasicUnit,	pvTmpRow, C_PrntBasicUnitPopup,	pvEndRow
				ggoSpread.SpreadUnLock	C_SafetyLT,			pvTmpRow, C_SafetyLT,				pvEndRow
				ggoSpread.SpreadUnLock	C_LossRate,			pvTmpRow, C_LossRate,				pvEndRow
				ggoSpread.SpreadUnLock	C_ValidToDt,		pvTmpRow, C_ValidToDt,			pvEndRow	
				ggoSpread.SpreadUnLock	C_Remark,			pvTmpRow, C_Remark,				pvEndRow
				
				If lgStrBOMHisFlg = "Y" Then		 
					ggoSpread.SpreadUnLock	C_ECNNo,		pvTmpRow, C_ECNNo,				pvEndRow
					ggoSpread.SpreadUnLock	C_ECNNoPopup,	pvTmpRow, C_ECNNoPopup,			pvEndRow
					ggoSpread.SSSetRequired	C_ECNNo,		pvTmpRow, pvEndRow		
				End If
				
				ggoSpread.SSSetRequired	C_ChildItemBaseQty,	pvTmpRow, pvEndRow
				ggoSpread.SSSetRequired C_ChildBasicUnit, 	pvTmpRow, pvEndRow
				ggoSpread.SSSetRequired C_PrntItemBaseQty, 	pvTmpRow, pvEndRow
				ggoSpread.SSSetRequired	C_PrntBasicUnit,	pvTmpRow, pvEndRow
				ggoSpread.SSSetRequired	C_ValidToDt,		pvTmpRow, pvEndRow
			End If
		End If	
	Else
		If Level = 0 Then	
			ggoSpread.SSSetProtected -1, pvStartRow, pvEndRow

			ggoSpread.SpreadUnLock	C_ChildItemCd,	pvStartRow, C_ChildItemPopup,	pvEndRow
			ggoSpread.SpreadUnLock	C_DrawingPath,	pvStartRow, C_DrawingPath,		pvEndRow
			ggoSpread.SSSetRequired C_ChildItemCd,	pvStartRow, pvEndRow

		Else
			ggoSpread.SSSetProtected -1, pvStartRow, pvEndRow
			
			If pvStartRow > 1 Then
				pvTmpRow = pvStartRow
			Else
				pvTmpRow = 2
			End If	
			
			ggoSpread.SpreadUnLock	C_ChildItemCd,		pvTmpRow, C_ChildItemPopup,		pvEndRow
			ggoSpread.SpreadUnLock	C_ChildItemBaseQty,	pvTmpRow, C_ChildItemBaseQty,		pvEndRow
			ggoSpread.SpreadUnLock	C_ChildBasicUnit,	pvTmpRow, C_ChildBasicUnitPopup,	pvEndRow
			ggoSpread.SpreadUnLock	C_PrntItemBaseQty,	pvTmpRow, C_PrntItemBaseQty,		pvEndRow
			ggoSpread.SpreadUnLock	C_PrntBasicUnit,	pvTmpRow, C_PrntBasicUnitPopup,	pvEndRow
			ggoSpread.SpreadUnLock	C_SafetyLT,			pvTmpRow, C_SafetyLT,				pvEndRow
			ggoSpread.SpreadUnLock	C_LossRate,			pvTmpRow, C_LossRate,				pvEndRow
			ggoSpread.SpreadUnLock	C_ValidFromDt,		pvTmpRow, C_ValidFromDt,			pvEndRow
			ggoSpread.SpreadUnLock	C_ValidToDt,		pvTmpRow, C_ValidToDt,			pvEndRow
			If lgStrBOMHisFlg = "Y" Then	 
				ggoSpread.SpreadUnLock	C_ECNNo,		pvTmpRow, C_ECNNo,				pvEndRow
				ggoSpread.SpreadUnLock	C_ECNNoPopup,	pvTmpRow, C_ECNNoPopup,			pvEndRow
				ggoSpread.SSSetRequired	C_ECNNo,		pvTmpRow, pvEndRow		
			End If
			ggoSpread.SpreadUnLock	C_Remark,			pvTmpRow, C_Remark,				pvEndRow
'			ggoSpread.SSSetRequired C_Seq, 				pvTmpRow, pvEndRow
			ggoSpread.SSSetRequired C_ChildItemCd, 		pvTmpRow, pvEndRow
			ggoSpread.SSSetRequired	C_ChildItemBaseQty,	pvTmpRow, pvEndRow
			ggoSpread.SSSetRequired C_ChildBasicUnit, 	pvTmpRow, pvEndRow
			ggoSpread.SSSetRequired C_PrntItemBaseQty, 	pvTmpRow, pvEndRow
			ggoSpread.SSSetRequired	C_PrntBasicUnit,	pvTmpRow, pvEndRow
			ggoSpread.SSSetRequired C_ValidFromDt, 		pvTmpRow, pvEndRow
			ggoSpread.SSSetRequired	C_ValidToDt,		pvTmpRow, pvEndRow
		End IF
	End If

    frm1.vspdData.ReDraw = True
	
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
			C_Level					= iCurColumnPos(1)
			C_Seq					= iCurColumnPos(2)
			C_ChildItemCd			= iCurColumnPos(3)
			C_ChildItemPopUp		= iCurColumnPos(4)
			C_ChildItemNm			= iCurColumnPos(5)
			C_Spec					= iCurColumnPos(6)
			C_ChildItemUnit			= iCurColumnPos(7)
			C_ItemAcct				= iCurColumnPos(8)
			C_ItemAcctNm			= iCurColumnPos(9)
			C_ProcType				= iCurColumnPos(10)
			C_ProcTypeNm			= iCurColumnPos(11)
			C_BomType				= iCurColumnPos(12)
			C_BomTypePopup			= iCurColumnPos(13)
			C_ChildItemBaseQty		= iCurColumnPos(14)
			C_ChildBasicUnit		= iCurColumnPos(15)
			C_ChildBasicUnitPopup	= iCurColumnPos(16)
			C_PrntItemBaseQty		= iCurColumnPos(17)
			C_PrntBasicUnit			= iCurColumnPos(18)
			C_PrntBasicUnitPopup	= iCurColumnPos(19)
			C_SafetyLT				= iCurColumnPos(20)
			C_LossRate				= iCurColumnPos(21)
			C_SupplyFlg				= iCurColumnPos(22)
			C_SupplyFlgNm			= iCurColumnPos(23)
			C_ValidFromDt			= iCurColumnPos(24)
			C_ValidToDt				= iCurColumnPos(25)
			C_ECNNo					= iCurColumnPos(26)
			C_ECNNoPopup			= iCurColumnPos(27)
			C_ECNDesc				= iCurColumnPos(28)
			C_ReasonCd				= iCurColumnPos(29)
			C_ReasonCdPopup			= iCurColumnPos(30)
			C_ReasonNm				= iCurColumnPos(31)
			C_DrawingPath			= iCurColumnPos(32)
			C_Remark				= iCurColumnPos(33)
			C_HdrItemCd				= iCurColumnPos(34)
			C_HdrBomNo				= iCurColumnPos(35)
			C_HdrProcType			= iCurColumnPos(36)
			C_ItemValidFromDt		= iCurColumnPos(37)
			C_ItemValidToDt			= iCurColumnPos(38)
			C_ItemAcctGrp			= iCurColumnPos(39)
			C_ReqTransNo			= iCurColumnPos(40)
			C_ReqTransDt			= iCurColumnPos(41)
			C_TransStatus			= iCurColumnPos(42)
			C_TransDt				= iCurColumnPos(43)
			C_Row					= iCurColumnPos(44)
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Level					= iCurColumnPos(1)
			C_Seq					= iCurColumnPos(2)
			C_ChildItemCd			= iCurColumnPos(3)
			C_ChildItemPopUp		= iCurColumnPos(4)
			C_ChildItemNm			= iCurColumnPos(5)
			C_Spec					= iCurColumnPos(6)
			C_ChildItemUnit			= iCurColumnPos(7)
			C_ItemAcct				= iCurColumnPos(8)
			C_ItemAcctNm			= iCurColumnPos(9)
			C_ProcType				= iCurColumnPos(10)
			C_ProcTypeNm			= iCurColumnPos(11)
			C_BomType				= iCurColumnPos(12)
			C_BomTypePopup			= iCurColumnPos(13)
			C_ChildItemBaseQty		= iCurColumnPos(14)
			C_ChildBasicUnit		= iCurColumnPos(15)
			C_ChildBasicUnitPopup	= iCurColumnPos(16)
			C_PrntItemBaseQty		= iCurColumnPos(17)
			C_PrntBasicUnit			= iCurColumnPos(18)
			C_PrntBasicUnitPopup	= iCurColumnPos(19)
			C_SafetyLT				= iCurColumnPos(20)
			C_LossRate				= iCurColumnPos(21)
			C_SupplyFlg				= iCurColumnPos(22)
			C_SupplyFlgNm			= iCurColumnPos(23)
			C_ValidFromDt			= iCurColumnPos(24)
			C_ValidToDt				= iCurColumnPos(25)
			C_ECNNo					= iCurColumnPos(26)
			C_ECNNoPopup			= iCurColumnPos(27)
			C_ECNDesc				= iCurColumnPos(28)
			C_ReasonCd				= iCurColumnPos(29)
			C_ReasonCdPopup			= iCurColumnPos(30)
			C_ReasonNm				= iCurColumnPos(31)
			C_DrawingPath			= iCurColumnPos(32)
			C_Remark				= iCurColumnPos(33)
			C_HdrItemCd				= iCurColumnPos(34)
			C_HdrBomNo				= iCurColumnPos(35)
			C_HdrProcType			= iCurColumnPos(36)
			C_ItemValidFromDt		= iCurColumnPos(37)
			C_ItemValidToDt			= iCurColumnPos(38)
			C_ItemAcctGrp			= iCurColumnPos(39)
			C_Row1					= iCurColumnPos(40)			
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
    Call InitSpreadSheet()      

	frm1.vspdData.ReDraw = False

    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData(2)

	Call SetSpreadColor(1, 1, 0, 1)
	
	With frm1
		.vspdData.Col = C_Row
		If .vspdData.Text <> "" Then
			For iIntCnt = 2 To .vspdData.MaxRows
				.vspdData.Col = C_HdrProcType
				.vspdData.Row = iIntCnt
	
				If UCase(Trim(.vspdData.Text)) = "O" Then
					Call SetFieldProp(iIntCnt, "D", "O")
				Else
					Call SetFieldProp(iIntCnt, "D", "P")
				End IF
			Next					
		End If
	End With	
	frm1.vspdData.ReDraw = True
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	Dim i, iStrArr, iStrNmArr
    Dim strCbo  
    Dim strCboCd
    Dim strCboNm 
	'****************************
    'List Minor code(유무상구분)
    '****************************
    'strCboCd = "" & vbTab & ""
    'strCboNm = "" & vbTab 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2201", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iStrArr = Split(lgF0, Chr(11))
    iStrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iStrArr) - 1
        strCboCd = strCboCd & iStrArr(i) & vbTab
        strCboNm = strCboNm & iStrNmArr(i) & vbTab
	Next

	iStrFree = iStrNmArr(1)
	
    ggoSpread.SetCombo strCboCd, C_SupplyFlg 'parent.ggoSpread.SSGetColsIndex()              'Supply Flag setting
    ggoSpread.SetCombo strCboNm, C_SupplyFlgNm 'parent.ggoSpread.SSGetColsIndex()            'Supply Flag Nm Setting
    
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex 
	
	If frm1.vspdData.MaxRows <= 1 Then Exit Sub
	If lngStartRow = 1 Then lngStartRow = 2
	
	With frm1.vspdData
	
		.ReDraw = False
		If frm1.hStatus.value = "N" Then		
			If lgStrBOMHisFlg = "Y" Then
				ggoSpread.SpreadUnLock	C_ECNNo,		lngStartRow, C_ECNNo, .MaxRows
				ggoSpread.SpreadUnLock	C_ECNNoPopup,	lngStartRow, C_ECNNoPopup, .MaxRows
				ggoSpread.SSSetRequired	C_ECNNo,		lngStartRow, .MaxRows
	
			Else
				ggoSpread.SSSetProtected C_ECNNo,		lngStartRow, .MaxRows
				ggoSpread.SSSetProtected C_ECNNoPopup,	lngStartRow, .MaxRows
	
			End If
		End If
		.ReDraw = True
	End With
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConBasePlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "기준공장팝업"											' 팝업 명칭 
	arrParam(1) = "B_PLANT A, P_PLANT_CONFIGURATION B"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBasePlantCd.value)							' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y'"		' Where Condition
	arrParam(5) = "기준공장"												' TextBox 명칭 
	
    arrField(0) = "A.PLANT_CD"												' Field명(0)
    arrField(1) = "A.PLANT_NM"												' Field명(1)
    
    arrHeader(0) = "공장"												' Header명(0)
    arrHeader(1) = "공장명"												' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConBasePlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
	
End Function


'------------------------------------------  OpenCondPlant2()  -------------------------------------------------
'	Name : OpenCondDestPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConDestPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "대상공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtDestPlantCd.value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "대상공장"					' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConDestPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtDestPlantCd.focus
	
End Function


'------------------------------------------  OpenReqTransNo()  -------------------------------------------------
'	Name : OpenReqTransNo()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenReqTransNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strPlantCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	strPlantCd = Trim(frm1.txtBasePlantCd.value)
	
	' 팝업 명칭 
	arrParam(0) = "이관의뢰번호"				
	' TABLE 명칭 
	arrParam(1) = "P_EBOM_TO_PBOM_MASTER A, B_ITEM B, B_PLANT C"						
	' Code Condition
	arrParam(2) = Trim(frm1.txtReqTransNo.value)	
	' Name Cindition
	arrParam(3) = ""							
	' Where Condition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD AND A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S")
	' TextBox 명칭 
	arrParam(5) = "이관의뢰번호"					

    arrField(0) = "A.REQ_TRANS_NO"				' Field명(0)	
    arrField(1) = "A.PLANT_CD"					' Field명(1)
    arrField(2) = "C.PLANT_NM"					' Field명(2)
    arrField(3) = "A.ITEM_CD"					' Field명(3)
    arrField(4) = "B.ITEM_NM"					' Field명(4)
    arrField(5) = "A.STATUS"					' Field명(5)
    
    arrHeader(0) = "이관의뢰번호"			' Header명(0)    
    arrHeader(1) = "대상공장"				' Header명(1)
    arrHeader(2) = "대상공장명"				' Header명(2)
    arrHeader(3) = "품목"					' Header명(3)
    arrHeader(4) = "품목명"					' Header명(4)
    arrHeader(5) = "이관상태"				' Header명(5)
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReqTransNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")  
	
	frm1.txtReqTransNo.focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal str, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If iPos = 0 Then
		If frm1.txtBasePlantCd.value = "" Then
			Call DisplayMsgBox("971012", "X", "기준공장", "X")
			frm1.txtBasePlantCd.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If		
	Else
		If frm1.txtDestPlantCd.value = "" Then
			Call DisplayMsgBox("971012", "X", "대상공장", "X")
			frm1.txtDestPlantCd.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If		
	End If
	
	IsOpenPop = True
	
	If iPos = 0 Then
		arrParam(0) = Trim(frm1.txtBasePlantCd.value)   ' Plant Code
	Else
		arrParam(0) = Trim(frm1.txtDestPlantCd.value)   ' Plant Code
	End If
		
	arrParam(1) = Trim(str)	' Item Code
	

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
    arrField(10) = 3	' Field명(1) : "SPECIFICATION"
    
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
		Call SetItemCd(arrRet, iPos)
	End If	
	
	If iPos = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtItemCd.focus
	Else
		Call SetActiveCell(frm1.vspdData,C_ChildItemCd,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	
End Function

'------------------------------------------  OpenECNNo()  -------------------------------------------------
'	Name : OpenECNNo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenECNNo(ByVal strECNNo, ByVal iPos)
	Dim arrRet
	Dim arrParam(4), arrField(3)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(strECNNo)   ' ECN No.

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
		Call SetECNNo(arrRet, iPos)
	End If
	
	Call SetActiveCell(frm1.vspdData,C_ECNNo,frm1.vspdData.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
		
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo(ByVal strItem, ByVal strBom, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtDestPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "대상공장", "X")
		frm1.txtDestPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True

	arrParam(0) = "BOM팝업"						' 팝업 명칭 
	arrParam(1) = "B_MINOR"							' TABLE 명칭 

	arrParam(2) = Trim(strBom)		' Code Condition
	
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	
	arrParam(5) = "BOM Type"						' TextBox 명칭 
	
    arrField(0) = "MINOR_CD"						' Field명(0)
    arrField(1) = "MINOR_NM"						' Field명(1)
        
    arrHeader(0) = "BOM Type"					' Header명(0)
    arrHeader(1) = "BOM 특성"					' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet, iPos)
	End If	
	
	If iPos = 0 Then
		Call SetFocusToDocument("M")
	Else
		Call SetActiveCell(frm1.vspdData,C_BomType,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End IF
	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenReasonCd()
'	Description : OpenReasonCd
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonCd(ByVal str, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iPos = 0 And UCase(frm1.txtReasonCd.className) = UCase(parent.UCN_PROTECTED) Then 
		IsOpenPop = False
		Exit Function
	End If		

	arrParam(0) = "변경근거팝업"
	arrParam(1) = "B_MINOR"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1402", "''", "S") & ""			
	arrParam(5) = "변경근거"			
	
    arrField(0) = "MINOR_CD"	
    arrField(1) = "MINOR_NM"	
   
    
    arrHeader(0) = "변경근거"		
    arrHeader(1) = "변경근거명"		
    
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetReason(arrRet, iPos)
	End If	
	
	If iPos = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtReasonCd.focus			
	Else
		Call SetActiveCell(frm1.vspdData,C_ReasonCd,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End IF
	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : OpenUnit
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit(ByVal str, ByVal iPos)
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
		Call SetUnit(arrRet, iPos)
	End If	
	
	If iPos = 0 Then
		Call SetActiveCell(frm1.vspdData,C_ChildBasicUnit,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement			
	Else
		Call SetActiveCell(frm1.vspdData,C_PrntBasicUnit,frm1.vspdData.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End IF
	
End Function

'------------------------------------------  SetConBasePlant()  ----------------------------------------------
'	Name : SetConBasePlant()
'	Description : Condition Base Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConBasePlant(byval arrRet)
	frm1.txtBasePlantCd.Value    = arrRet(0)		
	frm1.txtBasePlantNm.Value    = arrRet(1)		

	Call txtBasePlantCd_OnChange()
End Function

'------------------------------------------  SetConDestPlant()  ----------------------------------------------
'	Name : SetConDestPlant()
'	Description : Condition Destination Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConDestPlant(byval arrRet)
	frm1.txtDestPlantCd.Value    = arrRet(0)		
	frm1.txtDestPlantNm.Value    = arrRet(1)		

	Call txtDestPlantCd_OnChange()
End Function

'------------------------------------------  SetItemCd()  ----------------------------------------------
'	Name : SetItemCd()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byval arrRet, ByVal iPos)
	If iPos = 0 Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)
'		frm1.txtSpec.value		= arrRet(10)
'		frm1.txtBasicUnit.value = arrRet(4)
	Else 
		With frm1.vspdData
			.Col = C_ChildItemCd
			.Row = .ActiveRow
			.Text = arrRet(0)		
			
			Call LookUpItemByPlant(arrRet(0), .Row)

		End With
		
		lgBlnFlgChgValue = True	
	End IF
	
End Function
'------------------------------------------  SetECNNo()  --------------------------------------------------
'	Name : SetECNNo()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetECNNo(ByVal arrRet, ByVal iPos)
	If iPos = 0 Then
		frm1.txtECNNo.Value		= arrRet(0)
		frm1.txtECNDesc.Value	= arrRet(1)		
		frm1.txtReasonCd.Value	= arrRet(2)
		frm1.txtReasonNm.Value	= arrRet(3)		
				
	Else 
		With frm1.vspdData
			Call .SetText(C_ECNNo,		.ActiveRow, arrRet(0))
			Call .SetText(C_ECNDesc,	.ActiveRow, arrRet(1))
			Call .SetText(C_ReasonCd,	.ActiveRow, arrRet(2))
			Call .SetText(C_ReasonNm,	.ActiveRow, arrRet(3))
			
			ggoSpread.SpreadLock	C_ECNDesc,	-1, C_ReasonCdPopup
			
			Call vspdData_Change(1, .ActiveRow)		
		End With
	End IF
	
	lgBlnFlgChgValue = True	
	
End Function


'------------------------------------------  SetReason()  --------------------------------------------------
'	Name : SetReason()
'	Description : SetReason
'--------------------------------------------------------------------------------------------------------- 
Function SetReason(Byval arrRet, Byval iPos)

	With frm1.vspdData
		Call .SetText(C_ReasonCd, .ActiveRow, arrRet(0))
		Call .SetText(C_ReasonNm, .ActiveRow, arrRet(1))
	End With
	
End Function

'------------------------------------------  SetReqTransNo()  --------------------------------------------------
'	Name : SetReqTransNo()
'	Description : SetReqTransNo
'--------------------------------------------------------------------------------------------------------- 
Function SetReqTransNo(Byval arrRet)
	frm1.txtReqTransNo.Value	= arrRet(0)
	frm1.txtDestPlantCd.Value	= arrRet(1)		
	frm1.txtDestPlantNm.Value	= arrRet(2)
	frm1.txtItemCd.Value		= arrRet(3)	
	frm1.txtItemNm.Value		= arrRet(4)	
	frm1.hStatus.Value			= arrRet(5)	
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Of Measure Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(Byval arrRet, ByVal iPos)
	With frm1.vspdData
		Call .SetText(C_PrntBasicUnit,		.ActiveRow, arrRet(0))
		Call vspdData_Change(.Col, .Row)		' 변경이 일어났다고 알려줌 
	End With
	
	lgBlnFlgChgValue = True	
End Function

'==========================================================================================
'   Function Name :SetFieldProp
'   Function Desc :여러 Case에 따른 Field들의 속성을 변경한다.
'==========================================================================================

Function SetFieldProp(ByVal lRow, ByVal Level, ByVal ProcType)
	If lRow = 1 Then Exit Function
	ggoSpread.Source = frm1.vspdData
	If Level = "D" Then							'최상위품목이 아닌경우 
		ggoSpread.SSSetProtected	C_BomType,		lRow, lRow
		ggoSpread.SSSetProtected	C_BomTypePopup,	lRow, lRow
		
		If ProcType = "O" Then					'외주가공품인 경우 
			ggoSpread.SpreadUnLock	C_SupplyFlgNm,	lRow, C_SupplyFlgNm, lRow
			ggoSpread.SSSetRequired	C_SupplyFlgNm,	lRow, lRow
		ElseIf ProcType = "P" OR ProcType = "M" Then
			ggoSpread.SSSetProtected	C_SupplyFlgNm,	lRow, lRow
		End If
	Else
		ggoSpread.SpreadUnLock		C_BomType,		lRow, C_BomType, lRow													'최상위품목인 경우 
		ggoSpread.SSSetRequired		C_BomType,		lRow, lRow
		
		ggoSpread.SpreadUnLock		C_BomTypePopup,	lRow, C_BomTypePopup, lRow
		
	End If
	
End Function


Sub SetCookieVal()
	
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm")
	End If	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""

End Sub

'==========================================================================================
'   Function Name :LookUpItemByPlant
'   Function Desc :선택한 품목의 Item Acct를 읽는다.
'==========================================================================================
Sub LookUpItemByPlant(ByVal strItemCd, ByVal IRow)
    
    Err.Clear															'☜: Protect system from crashing

	Dim strSelect
	If strItemCd = "" Then Exit Sub
	
	frm1.vspdData.Col = C_ChildItemCd
	frm1.vspdData.Row = IRow		

	strSelect = " b.ITEM_NM, a.ITEM_ACCT, dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ", a.ITEM_ACCT) ITEM_ACCT_NM, a.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", a.PROCUR_TYPE) PROCUR_TYPE_NM, b.SPEC, b.BASIC_UNIT, dbo.ufn_GetItemAcctGrp(a.ITEM_ACCT) ITEM_ACCT_GRP "

	If 	CommonQueryRs2by2(strSelect, " B_ITEM_BY_PLANT a, B_ITEM b ", " a.ITEM_CD = b.ITEM_CD AND a.PLANT_CD = " & _
	    FilterVar(frm1.txtDestPlantCd.Value, "''", "S") & " AND a.ITEM_CD = " & FilterVar(strItemCd, "''", "S"), lgF0) = False Then
		Call DisplayMsgBox("122700", "X", strItemCd, "X")
		Call LookUpItemByPlantNotOk()
		Exit Sub
	End If

	lgF0 = Split(lgF0, Chr(11))

	Call LookUpItemByPlantOk(lgF0(1), lgF0(2), lgF0(3), lgF0(4), lgF0(5), lgF0(6), lgF0(7), IRow, lgF0(8))
End Sub

'==========================================================================================
'   Function Name :LookUpItemByPlantOk
'   Function Desc :선택한 품목의 존재여부를 Check함를 읽는다.
'==========================================================================================
Function LookUpItemByPlantOk(ByVal strItemNm, ByVal strItemAcct, ByVal strItemAcctNm, ByVal strProcType, ByVal strProcTypeNm, ByVal strSpec, ByVal strBasicUnit, ByVal IRow , ByVal strItemAcctGrp)

	Dim ChildItemCd
	Dim PrntItemCd
	Dim PrntBomNo
	Dim strLevel, strChildLevel, strComLevel

	isClicked =  False
	IsOpenPop = False

	With frm1.vspdData
		
		.ReDraw = False
		.Row = IRow		
		
		Call .GetText(C_ChildItemCd,	IRow, ChildItemCd)
		Call .GetText(C_HdrBomNo,		IRow, PrntBomNo)		'2003-09-08
		
		Call .SetText(C_ChildItemNm,	IRow, strItemNm)
		Call .SetText(C_ItemAcct,		IRow, strItemAcct)
		Call .SetText(C_ProcType,		IRow, strProcType)
		Call .SetText(C_ItemAcctNm,		IRow, strItemAcctNm)
		Call .SetText(C_ProcTypeNm,		IRow, strProcTypeNm)
		Call .SetText(C_Spec,			IRow, strSpec)
		Call .SetText(C_ChildItemUnit,	IRow, strBasicUnit)
		Call .SetText(C_ItemAcctGrp,	IRow, strItemAcctGrp)
		
		If IRow <> 1 Then					'자품목추가시 체크 로직 
			Call .GetText(C_HdrItemCd,		IRow, PrntItemCd)
			Call .GetText(C_HdrBomNo,		IRow, PrntBomNo)

			Call .SetText(C_ChildBasicUnit,		IRow, strBasicUnit)
			Call .SetText(C_PrntItemBaseQty,	IRow, "1")
			Call .SetText(C_ChildItemBaseQty,	IRow, "1")

			If UCase(ChildItemCd) = UCase(PrntItemCd) Then
				Call DisplayMsgBox("127421", "X", "모품목", "자품목")

				Call .SetText(C_ChildItemCd,	IRow, "")
				Call .SetText(C_ChildItemNm,	IRow, "")
				Call .SetText(C_ItemAcct,		IRow, "")
				Call .SetText(C_ProcType,		IRow, "")
				Call .SetText(C_ItemAcctNm,		IRow, "")
				Call .SetText(C_ProcTypeNm,		IRow, "")
				Call .SetText(C_Spec,			IRow, "")
				Call .SetText(C_ChildItemUnit,	IRow, "")
				Call .SetText(C_BomType,		IRow, "")
				Call .SetText(C_ItemAcctGrp,	IRow, "")
				
				Set gActiveElement = document.activeElement 
				Exit Function
			End If
			
			Call SetFieldProp(IRow, "D", "")					'Header:Create Detail:Create 
			Call .SetText(C_BomType,	IRow, PrntBomNo)
				
			If (UCase(Trim(strItemAcctGrp)) = "5GOODS" Or UCase(Trim(strItemAcctGrp)) = "6MRO") And PrntBomNo <> "E" Then
				Call DisplayMsgBox("182720", "X", "X", "X")
				Call .SetText(C_ChildItemCd,	IRow, "")
				Call .SetText(C_ChildItemNm,	IRow, "")
				Call .SetText(C_ItemAcct,		IRow, "")
				Call .SetText(C_ProcType,		IRow, "")
				Call .SetText(C_ItemAcctNm,		IRow, "")
				Call .SetText(C_ProcTypeNm,		IRow, "")
				Call .SetText(C_Spec,			IRow, "")
				Call .SetText(C_ChildItemUnit,	IRow, "")
				Call .SetText(C_BomType,		IRow, "")
				Call .SetText(C_ItemAcctGrp,	IRow, "")
				Exit Function 
			End If

		Else											'신규나 BOM복사시 체크 로직 
			If PrntBomNo <> "E" And Not (UCase(Trim(strItemAcctGrp)) = "1FINAL" Or UCase(Trim(strItemAcctGrp)) = "2SEMI" )Then
				Call DisplayMsgBox("182618", "X", "X", "X")
				Call .SetText(C_ChildItemCd,	IRow, "")
				Call .SetText(C_ChildItemNm,	IRow, "")
				Call .SetText(C_ItemAcct,		IRow, "")
				Call .SetText(C_ProcType,		IRow, "")
				Call .SetText(C_ItemAcctNm,		IRow, "")
				Call .SetText(C_ProcTypeNm,		IRow, "")
				Call .SetText(C_Spec,			IRow, "")
				Call .SetText(C_ChildItemUnit,	IRow, "")
				Call .SetText(C_BomType,		IRow, "")
				Call .SetText(C_ItemAcctGrp,	IRow, "")
				Exit Function 
			End If
		End If
				

		.Col = C_Level                                             '☜: Protect system from crashing
		strLevel = CLng(Replace(.Text, ".",""))
		strComLevel = strLevel + 1
			
		Do 
			.Col = C_Level
			.Row = .Row + 1
			If Trim(.Text) = "" Then
				strChildLevel = Clng(0)
			Else
				strChildLevel = Clng(Replace(Trim(.Text) , ".", ""))
			End If
			
			If (cstr(strChildLevel) = cstr(strComLevel)) Then
				.Col = C_HdrItemCd
				.Text = ChildItemCd
				.Col = C_HdrProcType
				.Text = strProcType
			End If 
						
		Loop While (strLevel < strChildLevel)
		
		.ReDraw = True
			
	End With

End Function

Function LookUpItemByPlantNotOk()
	
	Set gActiveElement = document.activeElement 	
	IRow = frm1.vspdData.Row
	With frm1.vspdData
	Call .SetText(C_ChildItemCd,	IRow, "")
	Call .SetText(C_ChildItemNm,	IRow, "")
	Call .SetText(C_ItemAcct,		IRow, "")
	Call .SetText(C_ProcType,		IRow, "")
	Call .SetText(C_ItemAcctNm,		IRow, "")
	Call .SetText(C_ProcTypeNm,		IRow, "")
	Call .SetText(C_Spec,			IRow, "")
	Call .SetText(C_ChildItemUnit,	IRow, "")
	Call .SetText(C_BomType,		IRow, "")
	Call .SetText(C_ItemAcctGrp,	IRow, "")
	.focus
	End With
	
    isClicked = False
    IsOpenPop = False
	
End Function

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

    Set gActiveSpdSheet = frm1.vspdData
    
    If Row <= 0 Or Col < 0 Then
		Exit Sub
	Else
		gMouseClickStatus = "SPC"
	    Call SetPopupMenuItemInf("0001110111")	
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

    Set gActiveSpdSheet = frm1.vspdData1
    
	gMouseClickStatus = "SP2C"
    Call SetPopupMenuItemInf("0000110111")

	If Row <= 0 Or Col < 0 Then
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
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
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

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
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
			Case C_SupplyFlgNm
				.Col = Col
				intIndex = .Value
				
				.Col = C_SupplyFlg
				.Value = intIndex
			
		End Select	
	End With	
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    isClicked =  True
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	
	With frm1.vspdData
		.Col = Col
		.Row = Row
				
		Select Case Col 
			Case C_ChildItemCd
				
				If .Text <> "" Then
					If CheckRunningBizProcess = True Then
					   Exit Sub
					End If 
					Call LookUpItemByPlant(Trim(.Text), Row)
				End If
				
			Case C_ECNNo

				Call LookupECN(.Text, 1)

			Case C_ReasonCd

				Call LookupReason(.Text, 1)

		End Select	
	End With

				
	If lgStrBOMHisFlg = "Y" And Row > 1 Then
		ggoSpread.SpreadUnLock	C_ECNNo,		Row, C_ECNNo,				Row
		ggoSpread.SpreadUnLock	C_ECNNoPopup,	Row, C_ECNNoPopup,			Row
		ggoSpread.SSSetRequired	C_ECNNo,		Row, Row
	End If

	isClicked = False
	
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : Spread sheet Button Click
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
    iCol = Col
    iRow = Row

	If Row <= 0 Then Exit Sub
	'----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
   
		If Col = C_BomTypePopup Then
			.Col = C_ChildItemCd
		    .Row = Row
		    
		    strTemp = .Text
		    
			.Col = C_BomType
		    .Row = Row
		    
		    Call OpenBomNo(strTemp, .Text, 1)
		    Call SetActiveCell(frm1.vspdData,C_BomType,Row,"M","X","X")
			Set gActiveElement = document.activeElement
			
		ElseIf Col = C_ChildBasicUnitPopup Then
			.Col = C_ChildBasicUnit
		    .Row = Row
		    
		    Call OpenUnit(.Text, 0)
		    Call SetActiveCell(frm1.vspdData,C_ChildBasicUnit,Row,"M","X","X")
			Set gActiveElement = document.activeElement
			
		ElseIf Col = C_PrntBasicUnitPopup Then
			.Col = C_PrntBasicUnit
		    .Row = Row
		    
		    Call OpenUnit(.Text, 1)
		    Call SetActiveCell(frm1.vspdData,C_PrntBasicUnit,Row,"M","X","X")
			Set gActiveElement = document.activeElement
			
		ElseIf Col = C_ChildItemPopup Then
			.Col = C_ChildItemCd
			.Row = Row

			If CheckRunningBizProcess = False Then
				Call OpenItemCd(.Text, 1)
				Call SetActiveCell(frm1.vspdData,C_ChildItemCd,Row,"M","X","X")
				Set gActiveElement = document.activeElement
			End If
			
		ElseIf Col = C_ECNNoPopup Then
			.Col = C_ECNNo
		    .Row = Row
		    
		    Call OpenECNNo(.Text, 1)
		    Call SetActiveCell(frm1.vspdData,C_ECNNo,Row,"M","X","X")
			Set gActiveElement = document.activeElement
			
		ElseIf Col = C_ReasonCdPopup Then
			.Col = C_ReasonCd
		    .Row = Row
		    
		    Call OpenReasonCd(.Text, 1)
		    Call SetActiveCell(frm1.vspdData,C_ReasonCd,Row,"M","X","X")
			Set gActiveElement = document.activeElement
 
		End If
    
    End With
End Sub

'==========================================================================================
'   Event Name : btnReq_OnClick()
'   Event Desc : 이관의뢰 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnRequest_OnClick()
    On Error Resume Next
    Err.Clear
	'이관의뢰상태를 'R'로 변경 
	
	frm1.hStatus.value = "R"
	frm1.hRequestingFlg.value = "Y"
	
	lgIntFlgMode = parent.OPMD_UMODE	'UPDATE로 변경 
	
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
   		frm1.hRequestingFlg.value = "N"
		Exit Sub
    End If  		

   	frm1.hRequestingFlg.value = "N"
	
End Sub

'==========================================================================================
'   Event Name : btnCancel_OnClick()
'   Event Desc : 이관의뢰취소 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnCancel_OnClick()
    On Error Resume Next
    Err.Clear
	'이관의뢰상태를 'R'->'N'으로 변경 
	
	If frm1.hStatus.value <> "R" Then
		Exit Sub
	End If
	
	frm1.hStatus.value = "N"
	
	frm1.hRequestingFlg.value = "Y"
	
	lgIntFlgMode = parent.OPMD_UMODE	'UPDATE로 변경 
	
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
   		frm1.hRequestingFlg.value = "N"
		Exit Sub
    End If  		

   	frm1.hRequestingFlg.value = "N"
	
End Sub

'==========================================================================================
'   Event Name : btnInit_OnClick()
'   Event Desc : 초기화 버튼을 클릭할 경우 발생 (기존 제조BOM과 동일하게 요청자료 자동생성))
'==========================================================================================
Sub btnInit_OnClick()

    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
	Dim i, j, strTemp
	Dim iStrDestPlantCd, iStrItemCd
    
    On Error Resume Next
    Err.Clear
  	
    '-----------------------
    'Check previous data area
    '(제조BOM에 수정된 내용이 있을경우)
    '-----------------------
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Sub
		End If
    End If

	iStrDestPlantCd = UCase(Trim(frm1.hDestPlantCd.value))
	iStrItemCd = UCase(Trim(frm1.hItemCd.value))    
    
    If iStrDestPlantCd = "" Or iStrItemCd = "" Then
		IntRetCD = DisplayMsgBox("800167", parent.VB_YES, "X", "X")
		Exit Sub
    End If
        
    Call ggoSpread.ClearSpreadData   
    
    With frm1
		strVal = BIZ_PGM_INIT_QRY_ID & "?txtMode=" & parent.UID_M0001	
		strVal = strVal & "&txtDestPlantCd=" & iStrDestPlantCd				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & iStrItemCd						'☆: 조회 조건 데이타 
		strVal = strVal & "&txtBomType=" & "1"
		strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
		
		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	
    End With
    
End Sub

'==========================================================================================
'   Event Name : btnCopy_OnClick()
'   Event Desc : 복사 버튼을 클릭할 경우 발생 (선택된설계BOM을 제조BOM에 추가)
'==========================================================================================
Sub btnCopy_OnClick()

    Dim iIntCnt
    Dim iStrFlag, iStrLevel
	Dim iSrcFromRow, iSrcToRow, iDestRow

    On Error Resume Next
    Err.Clear
	
    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData1.maxrows < 1 Then Exit Sub
  	
    With frm1    
		.vspdData.ReDraw = False
		
		iSrcFromRow =.vspdData1.SelBlockRow
		iSrcToRow =.vspdData1.SelBlockRow2
		
	    ggoSpread.Source = frm1.vspdData
	    
	    For i = iSrcFromRow To iSrcToRow

			    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
					Exit Sub
			    End If
	    
				Call FncInsertRow(1)

				iDestRow = .vspdData.MaxRows
				
    			.vspdData1.Row = i 

    			.vspdData1.Col = C_ChildItemCd 		: Call .vspdData.SetText(C_ChildItemCd		, iDestRow, .vspdData1.Text)
    			Call vspdData_Change(C_ChildItemCd , iDestRow )    			

    			.vspdData1.Col = C_ChildItemBaseQty : Call .vspdData.SetText(C_ChildItemBaseQty	, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_ChildBasicUnit 	: Call .vspdData.SetText(C_ChildBasicUnit	, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_PrntItemBaseQty	: Call .vspdData.SetText(C_PrntItemBaseQty	, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_PrntBasicUnit	: Call .vspdData.SetText(C_PrntBasicUnit	, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_SafetyLT			: Call .vspdData.SetText(C_SafetyLT			, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_LossRate			: Call .vspdData.SetText(C_LossRate			, iDestRow, .vspdData1.Text)

    			.vspdData1.Col = C_BomType 			: Call .vspdData.SetText(C_BomType			, iDestRow, "1")

    			.vspdData1.Col = C_ValidFromDt		: Call .vspdData.SetText(C_ValidFromDt		, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_ValidToDt		: Call .vspdData.SetText(C_ValidToDt		, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_ECNNo			: Call .vspdData.SetText(C_ECNNo			, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_ECNDesc			: Call .vspdData.SetText(C_ECNDesc			, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_ReasonCd			: Call .vspdData.SetText(C_ReasonCd			, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_ReasonNm			: Call .vspdData.SetText(C_ReasonNm			, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_DrawingPath		: Call .vspdData.SetText(C_DrawingPath		, iDestRow, .vspdData1.Text)
    			.vspdData1.Col = C_Remark			: Call .vspdData.SetText(C_Remark			, iDestRow, .vspdData1.Text)

		Next
		
		.vspdData.ReDraw = True
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
	'----------  Coding part  -------------------------------------------------------------   
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
        Exit Sub
    End If
	frm1.txtQueryType.value = "A"
	    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1, NewTop) Then
    
		If lgStrPrevKeyIndex1 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
        Exit Sub
    End If

	frm1.txtQueryType.value = "B"
	    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
    
		If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
End Sub

Sub txtBasePlantCd_OnChange()
	Dim arrVal, strWhere, strFrom
	
	strFrom = "B_PLANT A, P_PLANT_CONFIGURATION B"
	strWhere = 				" A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y' AND"
	strWhere = strWhere & 	" A.PLANT_CD = " & FilterVar(frm1.txtBasePlantCd.value, "''", "S")

	If Trim(frm1.txtBasePlantCd.value) <> "" Then
		If CommonQueryRs("A.PLANT_NM", strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrVal = Split(lgF0, Chr(11))
			frm1.txtBasePlantNm.Value = Trim(arrVal(0)) 
		Else
			frm1.txtBasePlantNm.Value = ""
		End If
	End If
End Sub
 
Sub txtDestPlantCd_OnChange()
	Dim arrVal
	
	ggoSpread.Source = frm1.vspdData
	
	If Trim(frm1.txtDestPlantCd.value) <> "" Then
		If CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtDestPlantCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrVal = Split(lgF0, Chr(11))
			frm1.txtDestPlantNm.Value = Trim(arrVal(0)) 
		Else
			frm1.txtDestPlantNm.Value = ""
		End If
		
		'BOM's History manage FLAG(Production Configuration)Check
		Call CommonQueryRs("BOM_HISTORY_FLG", "P_PLANT_CONFIGURATION", "PLANT_CD = " & FilterVar(frm1.txtDestPlantCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		If lgF0 = "" Or Left(lgF0, 1) = "N" Then
			lgStrBOMHisFlg = "N"
			'frm1.txtECNNo.value = ""
			'Call ggoOper.SetReqAttr(frm1.txtECNNo, "Q")
			ggoSpread.SpreadLock	C_ECNNo,	-1, C_ECNNoPopup
			ggoSpread.SpreadLock	C_ECNDesc,	-1, C_ECNDesc
			ggoSpread.SpreadLock	C_ReasonCd,	-1, C_ReasonCdPopup
		Else
			lgStrBOMHisFlg = "Y"
			ggoSpread.SpreadUnLock	C_ECNNo, -1, C_ECNNoPopup
			ggoSpread.SSSetRequired C_ECNNo, -1, -1

'			Call ggoOper.SetReqAttr(frm1.txtECNNo, "D")
		End If
	End If
End Sub

Sub LookupECN(ByVal strECNNo, ByVal iPos)
	Dim iArrECN(3)
	Dim iStrColSQL

	If Trim(strECNNo) <> "" Then
		iStrColSQL = "ECN_NO, ECN_DESC, REASON_CD, dbo.ufn_GetCodeName(" & FilterVar("P1402", "''", "S") & ", REASON_CD)"
		Call CommonQueryRs(iStrColSQL, "P_ECN_MASTER", "ECN_NO = " & FilterVar(strECNNo, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

		If Trim(lgF0) <> "" Then
			iArrECN(0) = Split(lgF0, Chr(11))(0)
			iArrECN(1) = Split(lgF1, Chr(11))(0)
			iArrECN(2) = Split(lgF2, Chr(11))(0)
			iArrECN(3) = Split(lgF3, Chr(11))(0)

			Call SetEcnNo(iArrECN, iPos)
			
			ggoSpread.SpreadLock	C_ReasonCd,	frm1.vspdData.ActiveRow, C_ReasonCdPopup, frm1.vspdData.ActiveRow
			ggoSpread.SpreadLock	C_ECNDesc,	frm1.vspdData.ActiveRow, C_ECNDesc, frm1.vspdData.ActiveRow
						
		Else
			Call frm1.vspdData.SetText(C_ECNDesc, frm1.vspdData.ActiveRow, "")
			Call frm1.vspdData.SetText(C_ReasonCd, frm1.vspdData.ActiveRow, "")
			Call frm1.vspdData.SetText(C_ReasonNm, frm1.vspdData.ActiveRow, "")
						
			ggoSpread.SpreadUnLock	C_ReasonCd,	frm1.vspdData.ActiveRow, C_ReasonCdPopup, frm1.vspdData.ActiveRow
			ggoSpread.SpreadUnLock	C_ECNDesc,	frm1.vspdData.ActiveRow, C_ECNDesc, frm1.vspdData.ActiveRow
			ggoSpread.SSSetRequired C_ReasonCd, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
			ggoSpread.SSSetRequired C_ECNDesc, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
		End If
	Else
		Call frm1.vspdData.SetText(C_ECNDesc, frm1.vspdData.ActiveRow, "")
		Call frm1.vspdData.SetText(C_ReasonCd, frm1.vspdData.ActiveRow, "")
		Call frm1.vspdData.SetText(C_ReasonNm, frm1.vspdData.ActiveRow, "")
		ggoSpread.SpreadLock	C_ReasonCd,	frm1.vspdData.ActiveRow, C_ReasonCdPopup
		ggoSpread.SpreadLock	C_ECNDesc,	frm1.vspdData.ActiveRow, C_ECNDesc
	End If
End Sub

Sub LookupReason(ByVal strReasonCd, ByVal iPos)
	Dim iArrReason(1)

	If Trim(strReasonCd) <> "" Then
		Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD = " & FilterVar("P1402", "''", "S") & " AND MINOR_CD = " & FilterVar(strReasonCd, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		If Trim(lgF0) <> "" Then
			iArrReason(0) = Split(lgF0, Chr(11))(0)
			iArrReason(1) = Split(lgF1, Chr(11))(0)
			Call SetReason(iArrReason, iPos)
		Else
			Call DisplayMsgBox("182803", "X", "X", "X")
			Call frm1.vspdData.SetText(C_ReasonCd, frm1.vspdData.ActiveRow, "")
			Call frm1.vspdData.SetText(C_ReasonNm, frm1.vspdData.ActiveRow, "")
			
			Call SetActiveCell(frm1.vspdData,C_ReasonCd,frm1.vspdData.ActiveRow,"M","X","X")
			Set gActiveElement = document.activeElement
			Exit Sub
		End If
	Else
		Call frm1.vspdData.SetText(C_ReasonCd, frm1.vspdData.ActiveRow, "")
		Call frm1.vspdData.SetText(C_ReasonNm, frm1.vspdData.ActiveRow, "")
	End If
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
    '(제조BOM에 수정된 내용이 있을경우)
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
    If frm1.txtBasePlantCd.value = "" Then
		frm1.txtBasePlantNm.value = ""
	End If
    If frm1.txtDestPlantCd.value = "" Then
		frm1.txtDestPlantNm.value = ""
	End If
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	'TAG='2XX'인 오브젝트 초기화 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    '스프레드 초기화 
    Call ggoSpread.ClearSpreadData
    
    Call InitVariables															'⊙: Initializes local global variables
    																			
    '-----------------------
    'Check condition area
    'TAG = '12'인 오브젝트에 값이 있는지 체크 
    '-----------------------
    If Not chkField(Document, "1") Then											'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
	frm1.txtQueryType.value = "*"
	    
    If DbQuery = False Then        											'☜: Query db data (설계BOM)   
		Exit Function           
    End If 
    
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
    slPlantCd = frm1.txtBasePlantCd.value
    slPlantNm = frm1.txtBasePlantNm.value
    
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    
    frm1.txtBasePlantCd.value = slPlantCd
    frm1.txtBasePlantNm.value = slplantNm
    
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call txtBasePlantCd_OnChange()
    
    Call SetToolbar("11101101000011")
    
    frm1.txtDestPlantCd.focus 
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
    If Not ggoSpread.SSDefaultCheck("Y") Then                                  '⊙: Delete시 Logic 변경(이진수)
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     				                                                  '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	Dim strLevel, strChildLevel
	Dim TempChildLevel
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_Level
		strLevel = Clng(Replace(.Text, ".", ""))
		
		Do 
			ggoSpread.EditUndo
			'.Col = C_ECNNo
			'Call LookupECN(.Text, 1)	'2003-09-13
			If .MaxRows = 0 Then Exit Do

			.Col = C_Level
			.Row = .ActiveRow
			If Trim(.Text) = "" Then
				strChildLevel = Clng(0)
			Else
				strChildLevel = Clng(Replace(Trim(.Text) , ".", ""))
			End If
		Loop While (strLevel < strChildLevel)
    End With
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim iIntReqRows, iIntCnt
	Dim iRow
	Dim strPrntLevel
	Dim strLevel
	Dim Level
	Dim PrntSeq
	Dim PrntItemCd
	Dim PrntBomNo
	Dim PrntItemAcct
	Dim PrntProcType
	Dim PrntItemAcctGrp
	Dim i, iInsertedRow
	Dim PrntBasicUnit

    On Error Resume Next
    Err.Clear

    FncInsertRow = False                                                         '☜: Processing is NG
	iIntReqRows = 1

	PrntBomNo = "1"	
	
	If Trim(frm1.txtDestPlantCd.value) = "" Then
		Call DisplayMsgBox("189220", "X", "X", "X")     
		Exit Function
	End If

       
    If Not chkField(Document, "2") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If

    With frm1
 
		.vspdData.Focus
		Set gActiveElement = document.activeElement
    
		ggoSpread.Source = .vspdData
    
		iRow = .vspdData.MaxRows
				
		Call .vspdData.GetText(C_Level, 1, strPrntLevel)

		If iRow >= 1 Then

			If IsNumeric(Trim(pvRowCnt)) Then
				iIntReqRows = CInt(pvRowCnt)
			Else
				iIntReqRows = AskSpdSheetAddRowCount()
				If iIntReqRows = "" Then
				    Exit Function
				End If
			End If
			
			Call .vspdData.GetText(C_Seq, .vspdData.MaxRows, PrntSeq)

			If Trim(PrntSeq) < Trim(.hMaxSeq.Value) Then
				PrntSeq = .hMaxSeq.Value
			End If
			Call .vspdData.GetText(C_ChildItemCd, 1, PrntItemCd)
			Call .vspdData.GetText(C_ItemAcct, 1, PrntItemAcct)
			Call .vspdData.GetText(C_ProcType, 1, PrntProcType)
			Call .vspdData.GetText(C_ItemAcctGrp, 1, PrntItemAcctGrp)

			Call .vspdData.GetText(C_ChildItemUnit, 1, PrntBasicUnit)
		
			strLevel = ".1"
			level = 1
		Else
			strLevel = "0"
		End If
		
        iInsertedRow = .vspdData.MaxRows + 1
        ggoSpread.InsertRow iInsertedRow, iIntReqRows
		
		.vspdData.EditMode = True
		.vspdData.ReDraw = False
		
		If lgIntFlgMode = parent.OPMD_CMODE And iRow < 1 Then
			Call SetSpreadColor(1, 1, 0, 0)
				
			.vspdData.Col = C_Level
			.vspdData.Text = strLevel
			Call .vspdData.SetText(C_Seq,.vspdData.ActiveRow, 0)
			Call .vspdData.SetText(C_BomType,.vspdData.ActiveRow, PrntBomNo)	'2003-09-08
			Call .vspdData.SetText(C_HdrBomNo,.vspdData.ActiveRow, PrntBomNo)	'2003-09-08
		Else
		
			For iIntCnt = iInsertedRow To iInsertedRow - 1 + iIntReqRows
				.vspdData.Row = iIntCnt

				Call .vspdData.SetText(C_Level,			iIntCnt, strLevel)
				Call .vspdData.SetText(C_Seq,			iIntCnt, PrntSeq + 1)
				Call .vspdData.SetText(C_PrntBasicUnit,	iIntCnt, PrntBasicUnit)
				Call .vspdData.SetText(C_HdrItemCd,		iIntCnt, PrntItemCd)
				Call .vspdData.SetText(C_HdrBomNo,		iIntCnt, PrntBomNo)
				Call .vspdData.SetText(C_BomType,		iIntCnt, PrntBomNo)
				Call .vspdData.SetText(C_HdrProcType,	iIntCnt, PrntProcType)
				Call .vspdData.SetText(C_SupplyFlg,		iIntCnt, "F")
				Call .vspdData.SetText(C_ValidFromDt,	iIntCnt, StartDate)
				Call .vspdData.SetText(C_ValidToDt,		iIntCnt, UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"))
				If lgStrBOMHisFlg = "Y" Then
					Call .vspdData.SetText(C_ECNNo,	iIntCnt, frm1.txtECNNo.value)
					Call .vspdData.SetText(C_ECNDesc,	iIntCnt, frm1.txtECNDesc.value)
					Call .vspdData.SetText(C_ReasonCd,	iIntCnt, frm1.txtReasonCd.value)
					Call .vspdData.SetText(C_ReasonNm,	iIntCnt, frm1.txtReasonNm.value)
				End If
				
			Next

			Call SetSpreadColor(iInsertedRow, iInsertedRow - 1 + iIntReqRows, Level, 0)
			
			For i = iInsertedRow To iInsertedRow - 1 + iIntReqRows
				Call .vspdData.SetText(C_SupplyFlgNm, i, iStrFree)
			Next
			
			If Trim(PrntProcType)= "O" Then					'상위품목이 외주가공품인 경우 
				ggoSpread.SpreadUnLock C_SupplyFlgNm,	iInsertedRow, C_SupplyFlgNm, iInsertedRow - 1 + iIntReqRows
				ggoSpread.SSSetRequired	C_SupplyFlgNm,	iInsertedRow, iInsertedRow - 1 + iIntReqRows
			Else
				ggoSpread.SSSetProtected C_SupplyFlgNm,	iInsertedRow, iInsertedRow - 1 + iIntReqRows
			End If
			
		End If

		.vspdData.ReDraw = True
		
	End With
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim iIntCnt
    Dim iStrFlag, iStrLevel

    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
	ggoSpread.Source = frm1.vspdData 

	For iIntCnt = frm1.vspdData.SelBlockRow  To frm1.vspdData.SelBlockRow2
		Call frm1.vspdData.GetText(C_Level, iIntCnt, iStrLevel)
		Call frm1.vspdData.GetText(0, iIntCnt, iStrFlag)
		If iStrFlag <> ggoSpread.InsertFlag And CInt(Replace(iStrLevel, ".", "")) = 1 Then
'			ggoSpread.EditUndo
			Call frm1.vspdData.SetText(0, iIntCnt, ggoSpread.DeleteFlag)
		End If
	Next
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
    Call parent.FncExport(parent.C_SINGLEMULTI)							'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)	                   '☜:화면 유형, Tab 유무 
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
' Function Name : DbQuery
' Function Desc : This function is data query and display(상단부)
'========================================================================================
Function DbQuery() 

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey
    Dim iStrBasePlantCd, iStrDestPlantCd, iStrItemCd, iStrReqTransNo
    Dim strQueryType

    DbQuery = False

    LayerShowHide(1)
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
	iStrBasePlantCd = UCase(Trim(frm1.txtBasePlantCd.value))
	iStrDestPlantCd = UCase(Trim(frm1.txtDestPlantCd.value))
	iStrItemCd = UCase(Trim(frm1.txtItemCd.value))
	iStrReqTransNo = UCase(Trim(frm1.txtReqTransNo.value))
	
	strQueryType = UCase(Trim(frm1.txtQueryType.value))						'☆: A : 설계BOM QUERY, B : 제조BOM QUERY, * : ALL

    With frm1
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001		
		strVal = strVal & "&QueryType=" & strQueryType	
		strVal = strVal & "&txtBasePlantCd=" & iStrBasePlantCd				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtDestPlantCd=" & iStrDestPlantCd				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemCd=" & iStrItemCd						'☆: 조회 조건 데이타 
		strVal = strVal & "&txtReqTransNo=" & iStrReqTransNo 				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtBaseDt=" & Trim(.txtBaseDt.value)
		strVal = strVal & "&txtSerchType=" & "1"
		strVal = strVal & "&txtBaseBomNo=" & "E"
		strVal = strVal & "&txtDestBomNo=" & "1"
		strVal = strVal & "&txtMaxRows="         & .vspdData.MaxRows
		strVal = strVal & "&txtMaxRows1="         & .vspdData1.MaxRows
        strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
        strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1         '☜: Next key tag
  
		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

    End With
    
    DbQuery = True
    'LayerShowHide(0)
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRow)										'☆: 조회 성공후 실행로직 
	
	Dim lRow
	Dim i
    '-----------------------
    'Reset variables area
    '-----------------------
    If frm1.vspdData.MaxRows > 0 Then
    	lgIntFlgMode = parent.OPMD_UMODE								'⊙: Indicates that current mode is Update mode
	End If
	
	Call ggoOper.LockField(Document, "Q")							'⊙: This function lock the suitable field

	frm1.hBasePlantCd.value = UCase(Trim(frm1.txtBasePlantCd.value))
	frm1.hDestPlantCd.value = UCase(Trim(frm1.txtDestPlantCd.value))
	frm1.hItemCd.value = UCase(Trim(frm1.txtItemCd.value))
		
	If frm1.vspdData.MaxRows < 1 Then
		Call SetToolbar("11101001000011")
	ElseIf frm1.vspdData.MaxRows = 1 Then
		Call SetToolbar("11111101000011")
	End If
	
	Call SetSpreadColor(1, 1, 0, 1)
	If frm1.vspdData.MaxRows > 1 Then
		Call SetSpreadColor(2, frm1.vspdData.MaxRows, 1, 1)
	End If
	
    frm1.vspdData.ReDraw = False
	With frm1
	
		.vspdData.Col = C_Row

		If .vspdData.Text <> "" Then

			For i = LngMaxRow To frm1.vspdData.MaxRows
				frm1.vspdData.Col = C_HdrProcType
				frm1.vspdData.Row = i

				If UCase(Trim(frm1.vspdData.Text)) = "O" Then
					Call SetFieldProp(i, "D", "O")
				Else
					Call SetFieldProp(i, "D", "P")
				End IF
				
			Next					
			
		End If
			
	End With
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
	
	Call txtDestPlantCd_OnChange()	'2003-08-11
	
	frm1.vspdData.focus
	lgBlnFlgChgValue = False
	frm1.vspdData.ReDraw = True

End Function
	
Function DbQueryNotOk()
    lgIntFlgMode = parent.OPMD_CMODE								'⊙: Indicates that current mode is Update mode
	
	Call SetToolbar("11100000000011") 
	  
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

	Dim IntRows 
    Dim strVal, strDel
	Dim strFromDate, strToDate
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
    
    LayerShowHide(1)

    With frm1
		.txtMode.value 			= parent.UID_M0002							'☜: 저장 상태 
		.txtFlgMode.value 		= lgIntFlgMode								'☜: 신규입력/수정 상태 
		.txtUpdtUserId.value 	= parent.gUsrID
		.txtInsrtUserId.value  	= parent.gUsrID
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
	
	
	frm1.vspdData.Row = 1
	frm1.vspdData.Col = C_ChildItemCd
	frm1.hItemCd.value = frm1.vspdData.Text
	frm1.vspdData.Col = C_BomType
	frm1.hBomType.value = frm1.vspdData.Text
	
	frm1.vspdData.Col = C_DrawingPath
	frm1.hDrawingPath.value = frm1.vspdData.Text

'	frm1.hHdrValidFromDt.value = StartDate
'	frm1.hHdrValidToDt.value = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	
	With frm1.vspdData
	    
	    For IntRows = 2 To .MaxRows
	    
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
					
			        .Col = C_Seq											'2
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_HdrItemCd										'3
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_HdrBomNo										'4
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_ChildItemCd									'5
			        strVal = strVal & Trim(.Text) & iColSep

			        .Col = C_BomType										'6
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_Level
			        
					If .Text <> "0" Then
						
						.Col = C_ChildItemBaseQty								'6
						If UNICDbl(.Text) = 0 Then
							Call DisplayMsgBox("970022", "X", "자품목기준수", "0")
							Set gActiveElement = document.activeElement
							Call LayerShowHide(0)
							Exit Function
						End If
						
					End If
					strVal = strVal & UNIConvNum(Trim(.Text), 1) & iColSep
					
					.Col = C_ChildBasicUnit									'7
			        strVal = strVal & Trim(.Text) & iColSep

					If .Text <> "0" Then
					
						.Col = C_PrntItemBaseQty								'8
						
						If UNICDbl(.Text) = 0 Then
							Call DisplayMsgBox("970022", "X", "모품목기준수", "0")	
							Set gActiveElement = document.activeElement 
							Call LayerShowHide(0)
							Exit Function
						End If

					End If
					
					strVal = strVal & UNIConvNum(Trim(.Text), 1) & iColSep
					
			        .Col = C_PrntBasicUnit									'9
			        strVal = strVal & Trim(.Text) & iColSep
	    
			        .Col = C_SafetyLT										'10
			        strVal = strVal & UNIConvNum(Trim(.Text), 0) & iColSep
	    			
	    			.Col = C_LossRate										'11
			        strVal = strVal & UNIConvNum(Trim(.Text), 0) & iColSep
			        
			        .Col = C_SupplyFlg										'12
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_ValidFromDt	
			        If Len(Trim(.Text)) Then
						If UNIConvDate(Trim(.Text)) = "" Then	 
							Call DisplayMsgBox("122116", "X", "X", "X")
							Set gActiveElement = document.activeElement 
							Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
							strFromDate = UNIConvDate(Trim(.Text))
						End If
					End If
			        
			        .Col = C_ValidToDt										'14
			        If Len(Trim(.Text)) Then
						If UNIConvDate(Trim(.Text)) = "" Then	 
							Call DisplayMsgBox("122116", "X", "X", "X")
							Set gActiveElement = document.activeElement 
							Call LayerShowHide(0)
							Exit Function
						Else
							strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
							strToDate = UNIConvDate(Trim(.Text))
						End If
					End If

					If strFromDate > strToDate Then
						Call DisplayMsgBox("972002", "X", "종료일", "시작일")
						Set gActiveElement = document.activeElement 
						Call LayerShowHide(0)
						Exit Function	
					End If	

			        .Col = C_ECNNo
			        strVal = strVal & UCase(Trim(.Text)) & iColSep
			        
			        .Col = C_ECNDesc
			        strVal = strVal & Trim(.Text) & iColSep
			        
			        .Col = C_ReasonCd
			        strVal = strVal & UCase(Trim(.Text)) & iColSep
	        
			        .Col = C_Remark											'15
			        strVal = strVal & Trim(.Text) & iColSep
		
			      	.Col = C_ReqTransNo
			      	strVal = strVal & UCase(Trim(.Text)) & iColSep
			      	
			      	.Col = C_ReqTransDt
			      	strVal = strVal & Trim(.Text) & iColSep
			      	
			      	.Col = C_TransStatus
			      	strVal = strVal & UCase(Trim(.Text)) & iColSep
			      	
			      	.Col = C_TransDt
			      	strVal = strVal & Trim(.Text) & parent.gRowSep  
			      	
			    Case ggoSpread.DeleteFlag

					strDel = ""
					
					strDel = strDel & "D" & iColSep	& IntRows & iColSep				'⊙: D=Delete
					
					.Col = C_Seq											'2
			        strDel = strDel & Trim(.Text) & iColSep
			        
					.Col = C_HdrItemCd										'3
			        strDel = strDel & Trim(.Text) & iColSep
			        
			        .Col = C_HdrBomNo										'4
			        strDel = strDel & Trim(.Text) & iColSep
			        
			        .Col = C_ChildItemCd									'5
			        strDel = strDel & Trim(.Text) & iColSep

			        .Col = C_BomType										'6
			        strDel = strDel & Trim(.Text) & iColSep
			        
					.Col = C_ChildItemBaseQty								'6
					strDel = strDel & UNIConvNum(Trim(.Text), 1) & iColSep
					
					.Col = C_ChildBasicUnit									'7
			        strDel = strDel & Trim(.Text) & iColSep

					.Col = C_PrntItemBaseQty								'8
					strDel = strDel & UNIConvNum(Trim(.Text), 1) & iColSep
					
			        .Col = C_PrntBasicUnit									'9
			        strDel = strDel & Trim(.Text) & iColSep
	    
			        .Col = C_SafetyLT										'10
			        strDel = strDel & UNIConvNum(Trim(.Text), 0) & iColSep
	    			
	    			.Col = C_LossRate										'11
			        strDel = strDel & UNIConvNum(Trim(.Text), 0) & iColSep
			        
			        .Col = C_SupplyFlg										'12
			        strDel = strDel & Trim(.Text) & iColSep
			        
			        .Col = C_ValidFromDt	
					strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep
			        
			        .Col = C_ValidToDt										'14
					strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep

			        .Col = C_ECNNo
			        strDel = strDel & UCase(Trim(.Text)) & iColSep

			        .Col = C_ECNDesc
			        strDel = strDel & Trim(.Text) & iColSep
			        
			        .Col = C_ReasonCd
			        strDel = strDel & UCase(Trim(.Text)) & iColSep
			        
			        .Col = C_Remark											'15
			        strDel = strDel & Trim(.Text) & iColSep
			        
			      	.Col = C_ReqTransNo
			      	strDel = strDel & UCase(Trim(.Text)) & iColSep 
			      	
			      	.Col = C_ReqTransDt
			      	strDel = strDel & Trim(.Text) & iColSep
			      	
			      	.Col = C_TransStatus
			      	strDel = strDel & UCase(Trim(.Text)) & iColSep
			      	
			      	.Col = C_TransDt
			      	strDel = strDel & Trim(.Text) & parent.gRowSep  

			End Select
			
		
			.Col = 0
			Select Case .Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			    
			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  		'한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")        '동적으로 한개의 form element를 동적으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT           ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
			       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                            '버퍼의 조정 증가치를 넘으면 
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
Function DbSaveOk()				
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

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
    strVal = strVal & "&txtBasePlantCd=" & Trim(frm1.hBasePlantCd.value)		'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtDestPlantCd=" & Trim(frm1.hDestPlantCd.value)		'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)				'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtReqTransNo=" & Trim(frm1.hReqTransNo.value)		'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtBomType=" & Trim(frm1.hBomType.value)		'☜: 삭제 조건 데이타								'☜: 삭제 조건 데이타 

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
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'=========================================================================================================
'	Name : LoadEBomHistory()
'	Description : 설계BOM이력조회 PAGE로 이동 
'========================================================================================================= 
Function LoadEBomHistory()

	PgmJump(EBOM_HISTORY_PGM_ID )

End Function

'=========================================================================================================
'	Name : LoadPBomCreate()
'	Description : 제조BOM등록(이관의뢰) PAGE로 이동 
'========================================================================================================= 
Function LoadPBomCreate()

	PgmJump(PBOM_CREATE_PGM_ID)

End Function

'=========================================================================================================
'	Name : LoadEBomToPBom()
'	Description : 제조BOM이관 PAGE로 이동 
'========================================================================================================= 
Function LoadEBomToPBom()

	PgmJump(EBOM_TO_PBOM_PGM_ID)

End Function

