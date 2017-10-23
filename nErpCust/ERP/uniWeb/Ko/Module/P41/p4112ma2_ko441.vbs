'******************************************  1.2 Global 변수/상수 선언  ***********************************
'   1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_INC_CLS_DT        = "p4100mb1.asp"                    '☆: LookUp Plant for Inventory Close Date
Const BIZ_PGM_QRY_ID            = "p4112mb21_ko441.asp"              '☆: List Production Order Header
Const BIZ_PGM_QRY2_ID	        = "p4112mb26_ko441.asp"				'☆: 비지니스 로직 ASP명
Const BIZ_PGM_SAVE_ID           = "p4112mb22_ko441.asp"              '☆: Manage Production Order Header
Const BIZ_PGM_SAVE_ID2          = "p4112mb27_ko441.asp"
Const BIZ_PGM_RELEASE_ID        = "p4112mb23_ko441.asp"              '☆: Release Production Order
Const BIZ_PGM_CANCEL_ID         = "p4112mb28_ko441.asp"              '☆: Release Production Order
Const BIZ_PGM_JUMPORDERRUN_ID   = "p4110ma1.asp"

' Grid 1(vspdData) - Operation
Dim C_ProdtOrderNo      '= 1
Dim C_ItemCode          '= 2
Dim C_ItemPopup         '= 3
Dim C_ItemName          '= 4
Dim C_Specification     '= 5
Dim C_OrderQty          '= 6
Dim C_OrderUnit         '= 7
Dim C_OrderUnitPopup    '= 8
Dim C_OrderQtyInBaseUnit'= 9
Dim C_BaseUnit          '= 10
Dim C_PlanStartDt       '= 11
Dim C_PlanEndDt         '= 12
Dim C_RoutingNo         '= 13
Dim C_RoutingNoPopup    '= 14
Dim C_SLCD              '= 15
Dim C_SLCDPopup         '= 16
Dim C_SLNM              '= 17
Dim C_WcCd              '= 18
Dim C_WcCdPopup         '= 19
Dim C_WcNm              '= 20
Dim C_OrderStatus       '= 21
Dim C_OrderStatusDesc   '= 22
Dim C_ReWorkFlag        '= 23
Dim C_Remark            '= 24
Dim C_BOMNo             '= 25
Dim C_OrderType         '= 26
Dim C_OrderTypeDesc     '= 27
Dim C_PlanOrderNo       '= 28
Dim C_TrackingNo        '= 29
Dim C_TrackingNoPopup   '= 30
Dim C_ScheduledStartDt  '= 31
Dim C_ScheduledEndDt    '= 32
Dim C_ValidFromDT       '= 33
Dim C_ValidToDT         '= 34
Dim C_OrderUnitMFG      '= 35
Dim C_OrderLtMFG        '= 36
Dim C_FixedMRPQty       '= 37
Dim C_MinMRPQty         '= 38
Dim C_MaxMRPQty         '= 39
Dim C_RoundQty          '= 40
Dim C_ScrapRateMFG      '= 41
Dim C_MPSMgr            '= 42
Dim C_MRPMgr            '= 43
Dim C_ProdMgr           '= 44
Dim C_ItemGroupCd       '= 45
Dim C_ItemGroupNm       '= 46
Dim C_MRPRunNo          '= 47
Dim C_ParentOrderNo     '= 48
Dim C_ParentOprNo       '= 49
Dim C_CostCd            '= 50
Dim C_CostPopup         '= 51
Dim C_CostNm            '= 52
Dim C_OprNo             '= 53

Dim C_BsItemCd          '= 54
Dim C_BsItemNm          '= 55
Dim C_BATCHCNT          '= 56       '20080307::HANC


' Grid 2(vspdData2) - Operation
Dim C_OprNo2         '= 1
Dim C_JobCd2         '= 2
Dim C_JobDesc2       '= 3
Dim C_WcCd2          '= 4
Dim C_WcNm2          '= 5
Dim C_ItemCd2        '= 6
Dim C_ItemCdPopup2   '= 7
Dim C_ItemNm2        '= 8
Dim C_Spec2          '= 9
Dim C_ReqQty2        '= 10
Dim C_BaseUnit2      '= 11
Dim C_IssuedQty2     '= 12
Dim C_ReqDt2         '= 13
Dim C_TrackingNo2    '= 14
Dim C_SlCd2          '= 15
Dim C_SlCdPopup2     '= 16
Dim C_SlNm2          '= 17
Dim C_ResvStatus2    '= 18
Dim C_ResvDesc2      '= 19
Dim C_IssueMthd2     '= 20
Dim C_IssueMthdDesc2 '= 21
Dim C_ReqNo2         '= 22
Dim C_Seq2           '= 23
Dim C_ProdtOrderNo2  '= 24
Dim C_OrderStatus2   '= 25
Dim C_InsideFlag2    '= 26
'==========================================  1.2.2 Global 변수 선언  =====================================
'   1. 변수 표준에 따름. prefix로 g를 사용함.
'   2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨
'=========================================================================================================
Dim lgInvCloseDt    '재고마감일
Dim lgCalType       'Calendar Type
Dim lgPlannedDate
Dim lgFlgQueryCnt
Dim whereVspdData
Dim lgSortKey1
Dim lgSortKey2
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop                       ' Popup
Dim lgAfterQryFlg
Dim lgButtonSelection
'#########################################################################################################
'                                               2. Function부
'
'   내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬
'   공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'                        2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'   Name : InitVariables()
'   Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgSortKey = 1

    lgSortKey1 = 1
    lgSortKey2 = 1

    whereVspdData = "TOP"

	lgButtonSelection = "Cancel"
    frm1.btnRelease.disabled = True
	frm1.btnRelease.value = "제조오더확정취소"
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'   Name : InitSpreadComboBox()
'   Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()

    Dim strCboCd
    Dim iCodeArr
    Dim iNameArr

    '****************************
    'List Minor code(Job Code)
    '****************************
    strCboCd =  "N" & vbTab & "Y"

    ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo strCboCd, C_ReWorkFlag

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1

    ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderTypeDesc

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc


    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobCd2
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobDesc2

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_OrderStatus2
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_OrderStatusDesc2
End Sub

'==========================================  2.2.6 InitData()  ==========================================
'   Name : InitData()
'   Description : Combo Display
'========================================================================================================
Sub InitData(ByVal lngStartRow)
    Dim intRow
    Dim intIndex

    With frm1.vspdData
        For intRow = lngStartRow To .MaxRows
            .Row = intRow
            .col = C_OrderStatus
            intIndex = .value
            .Col = C_OrderStatusDesc
            .value = intindex
            .Row = intRow
            .col = C_OrderType
            intIndex = .value
            .Col = C_OrderTypeDesc
            .value = intindex
        Next
    End With
End Sub

'==========================================  2.2.6 InitData()  ==========================================
'   Name : InitTrackingNCost()
'   Description : Enable/ Lock Tracking & Cost Center
'========================================================================================================
Sub InitTrackingNCost(ByVal strFieldProperty)
    With frm1.vspdData
        For LngRow = 1 To .MaxRows

            If strFieldProperty = "T" Or strFieldProperty = "A" Then
                .Row = LngRow
                .Col = C_TrackingNo

                If .Text = "*" Or .Text = "" Then
                    ggoSpread.SpreadLock C_TrackingNo, LngRow, C_TrackingNoPopup, LngRow
                    ggoSpread.SSSetProtected C_TrackingNo, LngRow, LngRow
                    ggoSpread.SSSetProtected C_TrackingNoPopup, LngRow, LngRow
                Else
                    ggoSpread.SpreadUnLock C_TrackingNo, LngRow, C_TrackingNoPopup, LngRow
                    ggoSpread.SSSetRequired C_TrackingNo, LngRow, LngRow
                End If

            End If

            If Ucase(Trim(frm1.hOprCostFlag.value)) = "Y" Then
                ggoSpread.SpreadUnLock C_CostCd, LngRow, C_CostPopup, LngRow
                ggoSpread.SSSetRequired C_CostCd, LngRow, LngRow
            Else
                ggoSpread.SpreadLock C_CostCd, LngRow, C_CostPopup, LngRow
                ggoSpread.SSSetProtected C_CostCd, LngRow, LngRow
                ggoSpread.SSSetProtected C_CostPopup, LngRow, LngRow
            End If

        Next
    End With

End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'   기능: 화면초기화
'   설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ==========================================
'   Name : SetDefaultVal()
'   Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
    frm1.txtProdFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -10, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
    frm1.txtProdToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 20, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.btnRelease.disabled = True
	frm1.btnRelease.value = "제조오더확정취소"
End Sub

'========================================  2.2.1 SetCookieVal()  ======================================
'   Name : SetCookieVal()
'   Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'===================================================================================================
Sub SetCookieVal()

    frm1.txtPlantCd.Value   = ReadCookie("txtPlantCd")
    frm1.txtPlantNm.value   = ReadCookie("txtPlantNm")
    frm1.txtProdOrderNo.Value   = ReadCookie("txtProdOrderNo")

    WriteCookie "txtPlantCd", ""
    WriteCookie "txtPlantNm", ""
    WriteCookie "txtProdOrderNo", ""
    WriteCookie "txtPGMID", ""

End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    Call InitSpreadPosVariables(pvSpdNo)

    '------------------------------------------
    ' Grid 1 - Operation Spread Setting
    '------------------------------------------
    If pvSpdNo = "A" Or pvSpdNo = "*" Then
        With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20050905", , Parent.gAllowDragDropSpread

        frm1.vspdData.ReDraw = False
        frm1.vspdData.MaxCols = C_BsItemNm + 1
        frm1.vspdData.MaxRows = 0

        Call GetSpreadColumnPos("A")

        ggoSpread.SSSetEdit     C_ProdtOrderNo, "제조오더번호", 18,,,18,2
        ggoSpread.SSSetEdit     C_ItemCode, "품목", 18,,,18,2
        ggoSpread.SSSetButton   C_ItemPopup
        ggoSpread.SSSetEdit     C_ItemName, "품목명", 25
        ggoSpread.SSSetEdit     C_Specification, "규격", 25
        ggoSpread.SSSetFloat    C_OrderQty,"오더수량",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_BATCHCNT,"배치수",15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit     C_OrderUnit, "오더단위", 8,,,3,2
        ggoSpread.SSSetButton   C_OrderUnitPopup
        ggoSpread.SSSetFloat    C_OrderQtyInBaseUnit, "기준수량",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit     C_BaseUnit, "기준단위", 8,,,3
        ggoSpread.SSSetEdit     C_SLCD, "창고", 10,,,7,2
        ggoSpread.SSSetButton   C_SLCDPopup
        ggoSpread.SSSetEdit     C_SLNM, "창고명", 12
        ggoSpread.SSSetEdit     C_WcCd,        "작업장", 10,,,10,2
        ggoSpread.SSSetButton   C_WcCdPopup
        ggoSpread.SSSetEdit     C_WcNm,        "작업장명", 20
        ggoSpread.SSSetEdit     C_RoutingNo, "라우팅", 10,,,7,2
        ggoSpread.SSSetButton   C_RoutingNoPopup
        ggoSpread.SSSetCombo    C_OrderStatus, "지시상태", 10
        ggoSpread.SSSetCombo    C_OrderStatusDesc, "지시상태", 10
        ggoSpread.SSSetCombo    C_ReWorkFlag, "재작업", 6
        ggoSpread.SSSetEdit     C_Remark, "비고", 20,,,20
        ggoSpread.SSSetEdit     C_BOMNo, "BOM Type", 10
        ggoSpread.SSSetCombo    C_OrderType, "지시구분", 10
        ggoSpread.SSSetCombo    C_OrderTypeDesc, "지시구분", 10
        ggoSpread.SSSetEdit     C_PlanOrderNo, "계획오더번호", 15
        ggoSpread.SSSetEdit     C_TrackingNo, "Tracking No.", 25,,,25,2
        ggoSpread.SSSetButton   C_TrackingNoPopup
        ggoSpread.SSSetDate     C_PlanStartDt, "착수예정일", 11, 2, parent.gDateFormat
        ggoSpread.SSSetDate     C_PlanEndDt, "완료예정일", 11, 2, parent.gDateFormat
        ggoSpread.SSSetDate     C_ScheduledStartDt, "착수계획일정", 11, 2, parent.gDateFormat
        ggoSpread.SSSetDate     C_ScheduledEndDt, "완료계획일정", 11, 2, parent.gDateFormat
        ggoSpread.SSSetDate     C_ValidFromDT, "품목유효일", 11, 2, parent.gDateFormat
        ggoSpread.SSSetDate     C_ValidToDT, "품목실효일", 11, 2, parent.gDateFormat
        ggoSpread.SSSetEdit     C_OrderUnitMFG, "오더단위", 10
        ggoSpread.SSSetEdit     C_OrderLtMFG, "제조 L/T", 10
        ggoSpread.SSSetFloat    C_FixedMRPQty, "고정오더수량",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_MinMRPQty, "최소오더수량",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_MaxMRPQty, "최대오더수량",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_RoundQty, "올림수",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_ScrapRateMFG, "제조품목불량률",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit     C_MPSMgr, "MPS 담당자", 10
        ggoSpread.SSSetEdit     C_MRPMgr, "MRP담당자", 10
        ggoSpread.SSSetEdit     C_ProdMgr, "생산담당자", 10
        ggoSpread.SSSetEdit     C_ItemGroupCd, "품목그룹",  15
        ggoSpread.SSSetEdit     C_ItemGroupNm, "품목그룹명", 30
        ggoSpread.SSSetEdit     C_MRPRunNo, "MRP Run번호", 15
        ggoSpread.SSSetEdit     C_ParentOrderNo, "상위오더번호", 18
        ggoSpread.SSSetEdit     C_ParentOprNo, "상위공정", 10
        ggoSpread.SSSetEdit     C_CostCd,   "작업지시 C/C", 10
        ggoSpread.SSSetButton   C_CostPopup
        ggoSpread.SSSetEdit     C_CostNm, "작업지시 C/C명", 14
        ggoSpread.SSSetEdit     C_OprNo,        "공정", 10

		ggoSpread.SSSetEdit     C_BsItemCd, "기준품목", 18,,,18,2
        ggoSpread.SSSetEdit     C_BsItemNm, "기준품목명", 25


        Call ggoSpread.MakePairsColumn(C_ItemCode,C_ItemPopup)
        Call ggoSpread.MakePairsColumn(C_TrackingNo, C_TrackingNoPopup)
        Call ggoSpread.MakePairsColumn(C_OrderUnit, C_OrderUnitPopup)
        Call ggoSpread.MakePairsColumn(C_SLCD, C_SLCDPopup)
        Call ggoSpread.MakePairsColumn(C_RoutingNo, C_RoutingNoPopup)
        Call ggoSpread.MakePairsColumn(C_CostCd, C_CostPopup)
        Call ggoSpread.MakePairsColumn(C_WcCd, C_WcCdPopup)

        Call ggoSpread.SSSetColHidden( C_OrderStatus, C_OrderStatus , True)

        Call ggoSpread.SSSetColHidden( C_WcCd, C_WcCd , True)                   '20080118::hanc
        Call ggoSpread.SSSetColHidden( C_WcCdPopup, C_WcCdPopup , True)         '20080118::hanc
        Call ggoSpread.SSSetColHidden( C_WcNm, C_WcNm , True)                   '20080118::hanc

        Call ggoSpread.SSSetColHidden( C_OrderType, C_OrderType , True)
        Call ggoSpread.SSSetColHidden( C_BOMNo, C_BOMNo , True)
        Call ggoSpread.SSSetColHidden( C_ValidFromDT, C_ProdMgr , True)
        Call ggoSpread.SSSetColHidden( C_ParentOrderNo, C_ParentOprNo , True)
        Call ggoSpread.SSSetColHidden( C_OprNo, C_OprNo , True)
        Call ggoSpread.SSSetColHidden( frm1.vspdData.MaxCols, frm1.vspdData.MaxCols , True)

        ggoSpread.SSSetSplit2(2)
        Call SetSpreadLock("A")

        frm1.vspdData.ReDraw = True

        End With
    End If

    '------------------------------------------
    ' Grid 2 - Component Spread Setting
    '------------------------------------------
    If pvSpdNo = "B" Or pvSpdNo = "*" Then

        With frm1.vspdData2

        ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20050905", , Parent.gAllowDragDropSpread
        frm1.vspdData2.ReDraw = False

        frm1.vspdData2.MaxCols = C_InsideFlag2 + 1
        frm1.vspdData2.MaxRows = 0

        Call GetSpreadColumnPos("B")

        ggoSpread.SSSetEdit     C_OprNo2,       "공정", 10
        ggoSpread.SSSetCombo    C_JobCd2,       "작업", 10
        ggoSpread.SSSetCombo    C_JobDesc2,      "작업명", 20
        ggoSpread.SSSetEdit     C_WcCd2,        "작업장", 10
        ggoSpread.SSSetEdit     C_WcNm2,        "작업장명", 20
        ggoSpread.SSSetEdit		C_ItemCd2,		"부품", 18,,,18,2
        ggoSpread.SSSetButton 	C_ItemCdPopup2
        ggoSpread.SSSetEdit		C_ItemNm2,		"부품명", 25
        ggoSpread.SSSetEdit		C_Spec2,			"규격", 25
        ggoSpread.SSSetFloat	C_ReqQty2, 		"필요수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit 	C_BaseUnit2, 		"단위", 7
        ggoSpread.SSSetFloat	C_IssuedQty2,	"출고수량", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
        ggoSpread.SSSetDate 	C_ReqDt2, 		"필요일", 11, 2, parent.gDateFormat
        ggoSpread.SSSetEdit 	C_TrackingNo2,	"Tracking No.", 25,,,25,2
        ggoSpread.SSSetEdit		C_SlCd2,	"출고창고", 10,,,7,2
        ggoSpread.SSSetButton 	C_SlCdPopup2
        ggoSpread.SSSetEdit		C_SlNm2,	"출고창고명", 20
        ggoSpread.SSSetEdit		C_ResvStatus2,	"출고상태", 10
        ggoSpread.SSSetEdit		C_ResvDesc2,	"출고상태", 10
        ggoSpread.SSSetEdit		C_IssueMthd2,	"출고방법", 10
        ggoSpread.SSSetEdit		C_IssueMthdDesc2,"출고방법", 10
        ggoSpread.SSSetEdit		C_ReqNo2,"", 10
        ggoSpread.SSSetEdit		C_Seq2,"", 10
        ggoSpread.SSSetEdit		C_ProdtOrderNo2,"", 10
        ggoSpread.SSSetEdit		C_OrderStatus2,"", 10
        ggoSpread.SSSetEdit		C_InsideFlag2,"사내/외", 10

        Call ggoSpread.MakePairsColumn(C_ItemCd2, C_ItemCdPopup2)
        Call ggoSpread.MakePairsColumn(C_SlCd2, C_SlCdPopup2)

        Call ggoSpread.SSSetColHidden( C_ResvStatus2, C_ResvStatus2, True)
        Call ggoSpread.SSSetColHidden( C_IssueMthd2, C_IssueMthd2, True)
        Call ggoSpread.SSSetColHidden( C_ReqNo2, C_ReqNo2, True)
        Call ggoSpread.SSSetColHidden( C_Seq2, C_Seq2, True)
        Call ggoSpread.SSSetColHidden( C_ProdtOrderNo2, C_ProdtOrderNo2, True)
        Call ggoSpread.SSSetColHidden( C_OrderStatus2, C_OrderStatus2, True)
        Call ggoSpread.SSSetColHidden( C_InsideFlag2, C_InsideFlag2, True)
        Call ggoSpread.SSSetColHidden( frm1.vspdData2.MaxCols, frm1.vspdData2.MaxCols , True)

'        ggoSpread.SSSetSplit2(5)
        Call SetSpreadLock("B")

        frm1.vspdData2.ReDraw = True

        End With
    End If


End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		If pvSpdNo = "A" Then
            '--------------------------------
            'Grid 1
            '--------------------------------
            ggoSpread.Source = frm1.vspdData

            frm1.vspdData.ReDraw = False
            ggoSpread.SpreadLock     C_ItemName, -1, C_ItemName
            ggoSpread.SpreadLock     C_Specification, -1, C_Specification
            ggoSpread.SpreadLock     C_SLNm, -1, C_SLNm
            ggoSpread.SpreadLock     C_OrderStatus, -1, C_OrderStatus
            ggoSpread.SpreadLock     C_OrderStatusDesc, -1, C_OrderStatusDesc
            ggoSpread.SpreadLock     C_BOMNo, -1, C_BOMNo
            ggoSpread.SpreadLock     C_OrderType, -1, C_OrderType
            ggoSpread.SpreadLock     C_OrderTypeDesc, -1, C_OrderTypeDesc
            ggoSpread.SpreadLock     C_PlanOrderNo, -1, C_PlanOrderNo
            ggoSpread.SpreadLock     C_ScheduledStartDt, -1, C_ScheduledStartDt
            ggoSpread.SpreadLock     C_ScheduledEndDt, -1, C_ScheduledEndDt
            ggoSpread.SpreadLock     C_OrderQtyInBaseUnit, -1, C_OrderQtyInBaseUnit
            ggoSpread.SpreadLock     C_BaseUnit, -1, C_BaseUnit
            ggoSpread.SpreadLock     C_MRPRunNo, -1, C_MRPRunNo
            ggoSpread.SpreadLock     C_ParentOrderNo, -1, C_ParentOprNo
            ggoSpread.SpreadLock     C_ItemGroupCd, -1, C_ItemGroupCd
            ggoSpread.SpreadLock     C_ItemGroupNm, -1, C_ItemGroupNm
            ggoSpread.SpreadLock     C_CostNm, -1, C_CostNm
            ggoSpread.SpreadLock     C_WcCd, -1, C_WcCd
            ggoSpread.SpreadLock     C_WcNm, -1, C_WcNm
            ggoSpread.SpreadLock     C_Remark, -1, C_Remark
            ggoSpread.SpreadLock     C_BsItemCd, -1, C_BsItemCd
            ggoSpread.SpreadLock     C_BsItemNm, -1, C_BsItemNm
            
            ggoSpread.SpreadLock     frm1.vspdData.MaxCols, -1, frm1.vspdData.MaxCols

            ggoSpread.SSSetRequired  C_OrderQty, -1
            ggoSpread.SSSetRequired  C_BATCHCNT, -1     '20080307
            
            ggoSpread.SSSetRequired  C_OrderUnit, -1
            ggoSpread.SSSetRequired  C_SLCd, -1
            ggoSpread.SSSetRequired  C_RoutingNo, -1
            ggoSpread.SSSetRequired  C_ReWorkFlag, -1
            ggoSpread.SSSetRequired  C_PlanStartDt, -1
            ggoSpread.SSSetRequired  C_PlanEndDt, -1

            frm1.vspdData.ReDraw = True
        End If

		If pvSpdNo = "B" Then
            '--------------------------------
            'Grid 2
            '--------------------------------
			ggoSpread.Source = frm1.vspdData2

			.vspdData2.ReDraw = False
			ggoSpread.SpreadLock C_OprNo2,         -1, C_OprNo2
			ggoSpread.SpreadLock C_JobCd2,         -1, C_JobCd2
			ggoSpread.SpreadLock C_JobDesc2,       -1, C_JobDesc2
			ggoSpread.SpreadLock C_WcCd2,          -1, C_WcCd2
			ggoSpread.SpreadLock C_WcNm2,          -1, C_WcNm2
			ggoSpread.SpreadLock C_ItemCdPopup2,   -1, C_ItemCdPopup2
			ggoSpread.SpreadLock C_ItemNm2,        -1, C_ItemNm2
			ggoSpread.SpreadLock C_Spec2,          -1, C_Spec2
			ggoSpread.SpreadLock C_BaseUnit2,      -1, C_BaseUnit2
			ggoSpread.SpreadLock C_IssuedQty2,     -1, C_IssuedQty2
			ggoSpread.SpreadLock C_TrackingNo2,    -1, C_TrackingNo2
			ggoSpread.SpreadLock C_SlCdPopup2,     -1, C_SlCdPopup2
			ggoSpread.SpreadLock C_SlNm2,          -1, C_SlNm2
			ggoSpread.SpreadLock C_ResvStatus2,    -1, C_ResvStatus2
			ggoSpread.SpreadLock C_ResvDesc2,      -1, C_ResvDesc2
			ggoSpread.SpreadLock C_IssueMthd2,     -1, C_IssueMthd2
			ggoSpread.SpreadLock C_IssueMthdDesc2, -1, C_IssueMthdDesc2
			ggoSpread.SpreadLock C_ReqNo2,         -1, C_ReqNo2
			ggoSpread.SpreadLock C_Seq2,           -1, C_Seq2
			ggoSpread.SpreadLock C_ProdtOrderNo2,  -1, C_ProdtOrderNo2
			ggoSpread.SpreadLock C_OrderStatus2,   -1, C_OrderStatus2
			ggoSpread.SpreadLock C_InsideFlag2,    -1, C_InsideFlag2

			ggoSpread.SpreadLock frm1.vspdData2.MaxCols, -1, frm1.vspdData2.MaxCols

			ggoSpread.SSSetRequired	 C_ItemCd2, -1
			ggoSpread.SSSetRequired  C_ReqQty2,	-1
			ggoSpread.SSSetRequired  C_ReqDt2, -1
			ggoSpread.SSSetRequired  C_SlCd2, -1

			frm1.vspdData2.Redraw = True
		End If
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1.vspdData

    .Redraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.SSSetRequired  C_ItemCode,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemName,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_Specification,       pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_OrderQty,            pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_BATCHCNT,            pvStartRow, pvEndRow    '20080307
    ggoSpread.SSSetRequired  C_OrderUnit,           pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderQtyInBaseUnit,  pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BaseUnit,            pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_SLCd,                pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SLNm,                pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_RoutingNo,           pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderStatus,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderStatusDesc,     pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_ReWorkFlag,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BOMNo,               pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderType,           pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_OrderTypeDesc,       pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_PlanOrderNo,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_TrackingNo,          pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_PlanStartDt,         pvStartRow, pvEndRow
    ggoSpread.SSSetRequired  C_PlanEndDt,           pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ScheduledStartDt,    pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ScheduledEndDt,      pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_MRPRunNo,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemGroupCd,         pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ItemGroupNm,         pvStartRow, pvEndRow

	ggoSpread.SSSetProtected C_BsItemCd,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_BsItemNm,			pvStartRow, pvEndRow

    If UCase(Trim(frm1.hOprCostFlag.value)) = "Y" Then
        ggoSpread.SSSetRequired  C_CostCd,          pvStartRow, pvEndRow
    Else
        ggoSpread.SSSetProtected C_CostCd,          pvStartRow, pvEndRow
    End If

    ggoSpread.SSSetProtected C_CostNm,              pvStartRow, pvEndRow
'20080227::hanc    ggoSpread.SSSetRequired C_WcCd,              pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_WcNm,              pvStartRow, pvEndRow

    .Col = 1
    .Row = .ActiveRow
    .Action = 0                         'parent.SS_ACTION_ACTIVE_CELL
    .EditMode = True

    .Redraw = True

    End With
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData2

		.Redraw = False
        ggoSpread.SSSetProtected C_OprNo2,         pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_JobCd2,         pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_JobDesc2,       pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_WcCd2,          pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_WcNm2,          pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_ItemCd2,         pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ItemNm2,        pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_Spec2,          pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_ReqQty2,         pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_BaseUnit2,      pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_IssuedQty2,     pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_ReqDt2,          pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_TrackingNo2,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_SlCd2,           pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_SlNm2,          pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ResvStatus2,    pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ResvDesc2,      pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_IssueMthd2,     pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_IssueMthdDesc2, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ReqNo2,         pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_Seq2,           pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ProdtOrderNo2,  pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_OrderStatus2,   pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_InsideFlag2,    pvStartRow, pvEndRow
		.Redraw = True

    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)

    ' Grid 1(vspdData) - Operation
    If pvSpdNo = "A" Or pvSpdNo = "*" Then
        C_ProdtOrderNo          = 1
        C_ItemCode              = 2
        C_ItemPopup             = 3
        C_ItemName              = 4
        C_Specification         = 5
        C_OrderQty              = 6
        C_BATCHCNT              = 7         '20080307::HANC
        C_OrderUnit             = 8 
        C_OrderUnitPopup        = 9 
        C_OrderQtyInBaseUnit    = 10
        C_BaseUnit              = 11
        C_PlanStartDt           = 12
        C_PlanEndDt             = 13
        C_RoutingNo             = 14
        C_RoutingNoPopup        = 15
        C_SLCD                  = 16
        C_SLCDPopup             = 17
        C_SLNM                  = 18
        C_WcCd                  = 19
        C_WcCdPopup             = 20
        C_WcNm                  = 21
        C_OrderStatus           = 22
        C_OrderStatusDesc       = 23
        C_ReWorkFlag            = 24
        C_Remark                = 25
        C_BOMNo                 = 26
        C_OrderType             = 27
        C_OrderTypeDesc         = 28
        C_PlanOrderNo           = 29
        C_TrackingNo            = 30
        C_TrackingNoPopup       = 31
        C_ScheduledStartDt      = 32
        C_ScheduledEndDt        = 33
        C_ValidFromDT           = 34
        C_ValidToDT             = 35
        C_OrderUnitMFG          = 36
        C_OrderLtMFG            = 37
        C_FixedMRPQty           = 38
        C_MinMRPQty             = 39
        C_MaxMRPQty             = 40
        C_RoundQty              = 41
        C_ScrapRateMFG          = 42
        C_MPSMgr                = 43
        C_MRPMgr                = 44
        C_ProdMgr               = 45
        C_ItemGroupCd           = 46
        C_ItemGroupNm           = 47
        C_MRPRunNo              = 48
        C_ParentOrderNo         = 49
        C_ParentOprNo           = 50
        C_CostCd                = 51
        C_Costpopup             = 52
        C_CostNm                = 53
        C_OprNo                 = 54
		C_BsItemCd              = 55
        C_BsItemNm              = 56

    End If

    ' Grid 2(vspdData2) - Operation
    If pvSpdNo = "B" Or pvSpdNo = "*" Then
        C_OprNo2         = 1
        C_JobCd2         = 2
        C_JobDesc2       = 3
        C_WcCd2          = 4
        C_WcNm2          = 5
        C_ItemCd2        = 6
        C_ItemCdPopup2   = 7
        C_ItemNm2        = 8
        C_Spec2          = 9
        C_ReqQty2        = 10
        C_BaseUnit2      = 11
        C_IssuedQty2     = 12
        C_ReqDt2         = 13
        C_TrackingNo2    = 14
        C_SlCd2          = 15
        C_SlCdPopup2     = 16
        C_SlNm2          = 17
        C_ResvStatus2    = 18
        C_ResvDesc2      = 19
        C_IssueMthd2     = 20
        C_IssueMthdDesc2 = 21
        C_ReqNo2         = 22
        C_Seq2           = 23
        C_ProdtOrderNo2  = 24
        C_OrderStatus2   = 25
        C_InsideFlag2    = 26
    End If
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
    Case "A"
        ggoSpread.Source = frm1.vspdData

        Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

        C_ProdtOrderNo          = iCurColumnPos(1)
        C_ItemCode              = iCurColumnPos(2)
        C_ItemPopup             = iCurColumnPos(3)
        C_ItemName              = iCurColumnPos(4)
        C_Specification         = iCurColumnPos(5)
        C_OrderQty              = iCurColumnPos(6)
        C_BATCHCNT              = iCurColumnPos(7)      '20080307::HANC
        C_OrderUnit             = iCurColumnPos(8) 
        C_OrderUnitPopup        = iCurColumnPos(9) 
        C_OrderQtyInBaseUnit    = iCurColumnPos(10)
        C_BaseUnit              = iCurColumnPos(11)
        C_PlanStartDt           = iCurColumnPos(12)
        C_PlanEndDt             = iCurColumnPos(13)
        C_RoutingNo             = iCurColumnPos(14)
        C_RoutingNoPopup        = iCurColumnPos(15)
        C_SLCD                  = iCurColumnPos(16)
        C_SLCDPopup             = iCurColumnPos(17)
        C_SLNM                  = iCurColumnPos(18)
        C_WcCd                  = iCurColumnPos(19)
        C_WcCdPopup             = iCurColumnPos(20)
        C_WcNm                  = iCurColumnPos(21)
        C_OrderStatus           = iCurColumnPos(22)
        C_OrderStatusDesc       = iCurColumnPos(23)
        C_ReWorkFlag            = iCurColumnPos(24)
        C_Remark                = iCurColumnPos(25)
        C_BOMNo                 = iCurColumnPos(26)
        C_OrderType             = iCurColumnPos(27)
        C_OrderTypeDesc         = iCurColumnPos(28)
        C_PlanOrderNo           = iCurColumnPos(29)
        C_TrackingNo            = iCurColumnPos(30)
        C_TrackingNoPopup       = iCurColumnPos(31)
        C_ScheduledStartDt      = iCurColumnPos(32)
        C_ScheduledEndDt        = iCurColumnPos(33)
        C_ValidFromDT           = iCurColumnPos(34)
        C_ValidToDT             = iCurColumnPos(35)
        C_OrderUnitMFG          = iCurColumnPos(36)
        C_OrderLtMFG            = iCurColumnPos(37)
        C_FixedMRPQty           = iCurColumnPos(38)
        C_MinMRPQty             = iCurColumnPos(39)
        C_MaxMRPQty             = iCurColumnPos(40)
        C_RoundQty              = iCurColumnPos(41)
        C_ScrapRateMFG          = iCurColumnPos(42)
        C_MPSMgr                = iCurColumnPos(43)
        C_MRPMgr                = iCurColumnPos(44)
        C_ProdMgr               = iCurColumnPos(45)
        C_ItemGroupCd           = iCurColumnPos(46)
        C_ItemGroupNm           = iCurColumnPos(47)
        C_MRPRunNo              = iCurColumnPos(48)
        C_ParentOrderNo         = iCurColumnPos(49)
        C_ParentOprNo           = iCurColumnPos(50)
        C_CostCd                = iCurColumnPos(51)
        C_Costpopup             = iCurColumnPos(52)
        C_CostNm                = iCurColumnPos(53)
        C_OprNo                 = iCurColumnPos(54)
		C_BsItemCd              = iCurColumnPos(55)
        C_BsItemNm              = iCurColumnPos(56)


    Case "B"
        ggoSpread.Source = frm1.vspdData2

        Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

        C_OprNo2         = iCurColumnPos(1)
        C_JobCd2         = iCurColumnPos(2)
        C_JobDesc2       = iCurColumnPos(3)
        C_WcCd2          = iCurColumnPos(4)
        C_WcNm2          = iCurColumnPos(5)
        C_ItemCd2        = iCurColumnPos(6)
        C_ItemCdPopup2   = iCurColumnPos(7)
        C_ItemNm2        = iCurColumnPos(8)
        C_Spec2          = iCurColumnPos(9)
        C_ReqQty2        = iCurColumnPos(10)
        C_BaseUnit2      = iCurColumnPos(11)
        C_IssuedQty2     = iCurColumnPos(12)
        C_ReqDt2         = iCurColumnPos(13)
        C_TrackingNo2    = iCurColumnPos(14)
        C_SlCd2          = iCurColumnPos(15)
        C_SlCdPopup2     = iCurColumnPos(16)
        C_SlNm2          = iCurColumnPos(17)
        C_ResvStatus2    = iCurColumnPos(18)
        C_ResvDesc2      = iCurColumnPos(19)
        C_IssueMthd2     = iCurColumnPos(20)
        C_IssueMthdDesc2 = iCurColumnPos(21)
        C_ReqNo2         = iCurColumnPos(22)
        C_Seq2           = iCurColumnPos(23)
        C_ProdtOrderNo2  = iCurColumnPos(24)
        C_OrderStatus2   = iCurColumnPos(25)
        C_InsideFlag2    = iCurColumnPos(26)
    End Select

End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'   기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다.
'         하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'------------------------------------------  OpenCondPlant()  -------------------------------------------
'   Name : OpenCondPlant()
'   Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    '2008-03-27 4:49오후 :: hanc
    IF Trim(frm1.txtPlantCd.Value) <> "" THEN
        Exit Function
    END IF

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "공장팝업"                    ' 팝업 명칭
    arrParam(1) = "B_PLANT"                         ' TABLE 명칭
    arrParam(2) = Trim(frm1.txtPlantCd.Value)       ' Code Condition
    arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)        ' Name Cindition
    arrParam(4) = ""                                ' Where Condition
    arrParam(5) = "공장"                        ' TextBox 명칭

    arrField(0) = "PLANT_CD"                        ' Field명(0)
    arrField(1) = "PLANT_NM"                        ' Field명(1)

    arrHeader(0) = "공장"                        ' Header명(0)
    arrHeader(1) = "공장명"                     ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetConPlant(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtPlantCd.focus

End Function

'------------------------------------------  OpenProdOrderNo()  ---------------------------------------
'   Name : OpenProdOrderNo()
'   Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()

    Dim arrRet
    Dim arrParam(8)
    Dim iCalledAspName

    If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
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
    arrParam(1) = frm1.txtProdFromDt.Text
    arrParam(2) = frm1.txtProdToDt.Text
    arrParam(3) = "OP"
    arrParam(4) = "OP"
    arrParam(5) = Trim(frm1.txtProdOrderNo.value)
    arrParam(6) = Trim(frm1.txtTrackingNo.value)
    arrParam(7) = Trim(frm1.txtItemCd.value)
    arrParam(8) = Trim(frm1.cboOrderType.value)

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetProdOrderNo(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtProdOrderNo.focus

End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'   Name : OpenTrackingInfo()
'   Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()

    Dim arrRet
    Dim arrParam(4)
    Dim iCalledAspName

    If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
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
    arrParam(3) = frm1.txtProdFromDt.Text
    arrParam(4) = frm1.txtProdToDt.Text

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetTrackingNo(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtTrackingNo.focus

End Function

'--------------------------------------  OpenTrackingInfo2()  ------------------------------------------
'   Name : OpenTrackingInfo2()
'   Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo2(Byval strCode, Byval Row)

    Dim arrRet
    Dim arrParam(4)
    Dim iCalledAspName

    If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
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
    arrParam(1) = strCode
    arrParam(2) = ""
    arrParam(3) = frm1.txtProdFromDt.Text
    arrParam(4) = frm1.txtProdToDt.Text

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetTrackingNo2(arrRet, Row)
    End If

End Function

'------------------------------------------  OpenCondPlant()  -------------------------------------------
'   Name : OpenMRPRunNo()
'   Description : Condition MRP Run No. PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenMRPRunNo()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True

    arrParam(0) = "MRP Run번호 팝업"                    ' 팝업 명칭
    arrParam(1) = "(SELECT DISTINCT ORDER_NO A, CONFIRM_DT B, " & FilterVar("제조오더전개", "''", "S") & " C FROM P_EXPL_HISTORY WHERE ISNULL(CONFIRM_DT, '') <> '' AND PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "
    arrParam(1) = arrParam(1) & "UNION SELECT DISTINCT RUN_NO A, START_DT B, " & FilterVar("MRP전개", "''", "S") & " C FROM P_MRP_HISTORY WHERE PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ") AS a"
    arrParam(2) = Trim(frm1.txtMRPRunNo.Value)      ' Code Condition
    arrParam(3) = ""
    arrParam(4) = "" ' Where Condition
    arrParam(5) = "MRP Run번호"                     ' TextBox 명칭

    arrField(0) = "A"                       ' Field명(0)
    arrField(1) = "B"   ' Field명(1)
    arrField(2) = "C"                       ' Field명(2)

    arrHeader(0) = "RUN NO."                         ' Header명(0)
    arrHeader(1) = "일자"                       ' Header명(1)
    arrHeader(2) = "전개구분"                       ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetMRPRunNo(arrRet)
    End If

End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'   Name : OpenItemInfo()
'   Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode)

    Dim arrRet
    Dim arrParam(5), arrField(6)
    Dim iCalledAspName

    If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
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

    arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
    arrParam(1) = strCode                       ' Item Code
    arrParam(2) = "12!MO"                       ' Combo Set Data:"1029!MP" -- 품목계정 구분자 조달구분
    arrParam(3) = ""                            ' Default Value

    arrField(0) = 1 '"ITEM_CD"                  ' Field명(0)
    arrField(1) = 2 '"ITEM_NM"                  ' Field명(1)

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
                  "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetItemInfo(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtItemCd.focus

End Function

'===========================================================================================================
Function OpenItemGroup()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

    IsOpenPop = True

    arrParam(0) = "품목그룹팝업"
    arrParam(1) = "B_ITEM_GROUP"
    arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
    arrParam(3) = ""
    arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
    arrParam(5) = "품목그룹"

    arrField(0) = "ITEM_GROUP_CD"
    arrField(1) = "ITEM_GROUP_NM"

    arrHeader(0) = "품목그룹"
    arrHeader(1) = "품목그룹명"

    arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetItemGroup(arrRet)
    End If

    Call SetFocusToDocument("M")
    frm1.txtItemGroupCd.focus

End Function

'------------------------------------------  OpenItemInfo2()  -------------------------------------------------
'   Name : OpenItemInfo2()
'   Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo2(Byval strCode, Byval Row)

    Dim arrRet
    Dim arrParam(5), arrField(17)
    Dim iCalledAspName

    If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
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

    arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
    arrParam(1) = strCode                       ' Item Code
    arrParam(2) = "12!MO"                       ' Combo Set Data:"1029!MP" -- 품목계정 구분자 조달구분
    arrParam(3) = ""                            ' Default Value

    arrField(0) = 1                             'ITEM_CD
    arrField(1) = 2                             'ITEM_NM
    arrField(2) = 4                             'BASIC_UNIT
    arrField(3) = 28                            'ORDER_LT
    arrField(4) = 33                            'MIN_MRP_QTY
    arrField(5) = 34                            'MAX_MRP_QTY
    arrField(6) = 35                            'ROND_QTY
    arrField(7) = 37                            'MPS_FLAG
    arrField(8) = 25                            'Tracking Flag
    arrField(9) = 26                            'UNIT_OF_ORDER
    arrField(10) = 15                           'MAJOR_SL_CD
    arrField(11) = 13                           'PHANTOM_FLG
    arrField(12) = 9                            'PROCUR_TYPE
    arrField(13) = 32                           'FIXED_MRP_QTY
    arrField(14) = 18                           'VALID_FROM_DT
    arrField(15) = 19                           'VALID_TO_DT
    arrField(16) = 17                           'VALID_FLG
    arrField(17) = 3                            'SPEC

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
                  "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        Exit Function
    Else
        frm1.vspdData.Row = Row
        frm1.vspdData.Col = C_ItemCode
        frm1.vspdData.Text = arrRet(0)
        Call LookUpItemByPlant(arrRet(0), Row)
    End If

End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'   Name : OpenSLCd()
'   Description : Storage Location PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSLCd(Byval strCode, Byval Row)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    IsOpenPop = True

    arrParam(0) = "창고팝업"                                            ' 팝업 명칭
    arrParam(1) = "B_STORAGE_LOCATION"                                      ' TABLE 명칭
    arrParam(2) = strCode                                                   ' Code Condition
    arrParam(3) = ""'Trim(frm1.txtSLNm.Value)                               ' Name Cindition
    arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")    ' Where Condition
    arrParam(5) = "창고"                                                ' TextBox 명칭
    arrField(0) = "SL_CD"                                                   ' Field명(0)
    arrField(1) = "SL_NM"                                                   ' Field명(1)
    arrHeader(0) = "창고"                                               ' Header명(0)
    arrHeader(1) = "창고명"                                             ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetSLCd(arrRet, Row)
    End If

End Function

'====================  OpenRoutingNo  ======================================
' Function Name : OpenRoutingNo
' Function Desc : OpenRoutingNo Reference Popup
'===========================================================================
Function OpenRoutingNo(Byval strRouting, Byval Row)

    Dim arrRet
    Dim arrParam(6), arrField(6), arrHeader(6)
    Dim strItemCode

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    frm1.vspdData.Row = Row
    frm1.vspdData.Col = C_ItemCode
    strItemCode = frm1.vspdData.Value
    If frm1.vspdData.Value = "" Then
        Call DisplayMsgBox("971012","X", "품목","X")
        IsOpenPop = False
        Exit Function
    End If

    arrParam(0) = "라우팅 팝업"                 ' 팝업 명칭
    arrParam(1) = "P_ROUTING_HEADER"                ' TABLE 명칭
    arrParam(2) = strRouting                        ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
                    "AND P_ROUTING_HEADER.ITEM_CD = " & FilterVar(UCase(strItemCode), "''", "S")
    arrParam(5) = "라우팅"

    arrField(0) = "ROUT_NO"                         ' Field명(0)
    arrField(1) = "DESCRIPTION"                     ' Field명(1)
    arrField(2) = "BOM_NO"                          ' Field명(1)
    arrField(3) = "MAJOR_FLG"                       ' Field명(1)

    arrHeader(0) = "라우팅"                     ' Header명(0)
    arrHeader(1) = "라우팅명"                   ' Header명(1)
    arrHeader(2) = "BOM Type"                   ' Header명(1)
    arrHeader(3) = "주라우팅"                   ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) <> "" Then
        Call SetRoutingNo(arrRet, Row)
    End If

End Function

 '------------------------------------------  OpenUnit()  -------------------------------------------------
'   Name : OpenUnit()
'   Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit(Byval strUnit, Byval Row)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    arrParam(0) = "단위팝업"
    arrParam(1) = "B_UNIT_OF_MEASURE"
    arrParam(2) = Trim(strUnit)
    arrParam(3) = ""
    arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & ""
    arrParam(5) = "단위"

    arrField(0) = "UNIT"
    arrField(1) = "UNIT_NM"

    arrHeader(0) = "단위"
    arrHeader(1) = "단위명"

    IsOpenPop = True

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetUnit(arrRet, Row)
    End If

End Function

'------------------------------------------  OpenCostCtr()  ----------------------------------------------
'   Name : OpenCostCtr()
'   Description : Cost Center Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCostCtr(ByVal StrCostCd, ByVal Row)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True  Then Exit Function

    If Trim(frm1.txtPlantCd.value) = "" Then
        Call DisplayMsgBox("971012", "X", "공장", "X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End If

    IsOpenPop = True

    arrParam(0) = "Cost Center 팝업"            ' 팝업 명칭
    arrParam(1) = "B_COST_CENTER"                   ' TABLE 명칭
    arrParam(2) = Trim(StrCostCd)       ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = "B_COST_CENTER.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
                " AND B_COST_CENTER.COST_TYPE ='M'" & _
                " AND B_COST_CENTER.DI_FG ='D'"         ' Where Condition
    arrParam(5) = "Cost Center"                 ' TextBox 명칭

    arrField(0) = "COST_CD"                         ' Field명(0)
    arrField(1) = "COST_NM"                         ' Field명(1)

    arrHeader(0) = "Cost Center"                ' Header명(0)
    arrHeader(1) = "Cost Center 명"             ' Header명(1)

    arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetCostCtr(arrRet, Row)
    End If

End Function


'------------------------------------------  OpenPartRef()  -------------------------------------------------
'   Name : OpenPartRef()
'   Description : Part Reference PopUp
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
        Call DisplayMsgBox("971012","X", "공장","X")
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

    arrParam(0) = Trim(frm1.txtPlantCd.value)       '☆: 조회 조건 데이타

    With frm1.vspdData
        If .MaxRows <= 0 Then Exit Function
        .Row = .ActiveRow
        .Col = C_ProdtOrderNo
        arrParam(1) = .Text
    End With

    IsOpenPop = True

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent,arrParam(0), arrParam(1), arrParam(2)), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

End Function

'------------------------------------------  OpenOprRef()  -------------------------------------------
'   Name : OpenOprRef()
'   Description : Operation Reference PopUp
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
        Call DisplayMsgBox("971012","X", "공장","X")
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

    arrParam(0) = Trim(frm1.txtPlantCd.value)       '☆: 조회 조건 데이타

    With frm1.vspdData
        If .MaxRows <= 0 Then Exit Function
        .Row = .ActiveRow
        .Col = C_ProdtOrderNo
        arrParam(1) = .Text
    End With

    IsOpenPop = True

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
        "dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

End Function

'------------------------------------------  OpenStockRef()  -------------------------------------------
'   Name : OpenStockRef()
'   Description : Stock Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenStockRef()
    Dim arrRet
    Dim arrParam(5)
    Dim iCalledAspName

    If IsOpenPop = True Then Exit Function

    If lgIntFlgMode = parent.OPMD_CMODE Then
        Call DisplayMsgBox("900002", "x", "x", "x")
        Exit Function
    End If

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    iCalledAspName = AskPRAspName("P4212RA1")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4212RA1", "X")
        IsOpenPop = False
        Exit Function
    End If

    With frm1.vspdData
        If .MaxRows <= 0 Then Exit Function
        .Row = .ActiveRow
        .Col = C_ItemCode

        If .text = "" Then
            Call DisplayMsgBox("971012","X", "품목","X")
            .focus
            .Row = .ActiveRow
            .Col = C_ItemCode
            .Action = 0
            .SelStart = 0
            Set gActiveElement = document.activeElement
            IsOpenPop = False
            Exit Function
        End If

        arrParam(0) = Trim(UCase(frm1.txtPlantCd.value))
        arrParam(1) = Trim(UCase(.text))
        .Col = C_ItemName
        arrParam(2) = Trim(.text)
        .Col = C_SLCD
        arrParam(3) = Trim(UCase(.text))
        .Col = C_SLNM
        arrParam(4) = Trim(.text)

    End With

    IsOpenPop = True

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4)), _
        "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

End Function

'------------------------------------------  OpenConWC2()  -------------------------------------------------
'	Name : OpenConWC2()
'	Description : Condition Work Center PopUp for Grid 2
'---------------------------------------------------------------------------------------------------------
Function OpenConWC2(Byval strCode, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"											' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = ""'strName												' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")	' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConWC2(arrRet,Row)
	End If	
	
End Function

'------------------------------------------  SetItemInfo()  -------------------------------------------
'   Name : SetItemInfo()
'   Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
    With frm1
        .txtItemCd.value = arrRet(0)
        .txtItemNm.value = arrRet(1)
    End With
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
    frm1.txtItemGroupCd.Value    = arrRet(0)
    frm1.txtItemGroupNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetItemInfo2()  -------------------------------------------
'   Name : SetItemInfo2()
'   Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo2(Byval arrRet, Byval Row)

    With frm1

        ggoSpread.Source = frm1.vspdData

        If arrRet(11) = "Y" Then 'PHANTOM_FLG
            Call DisplayMsgBox("189214", "x", "x", "x")
            Exit Function
        End If

        If arrRet(16) = "N" Then 'VALID_FLG
            Call DisplayMsgBox("122729", "x", "x", "x")
            Exit Function
        End If

        If arrRet(8) <> "Y" Then
            ggoSpread.SpreadLock C_TrackingNo, Row, C_TrackingNoPopup, Row
            ggoSpread.SSSetProtected C_TrackingNo, Row, Row
            ggoSpread.SSSetProtected C_TrackingNoPopup, Row, Row
            .vspdData.Row = Row
            .vspdData.Col = C_TrackingNo
            .vspdData.Text = "*"
        Else
            ggoSpread.SpreadUnLock C_TrackingNo, Row, C_TrackingNoPopup, Row
            ggoSpread.SSSetRequired C_TrackingNo, Row, Row
            .vspdData.Row = Row
            .vspdData.Col = C_TrackingNo
            .vspdData.Text = ""
        End If

        .vspdData.Row = Row

        .vspdData.Col = C_ItemCode
        .vspdData.Text = arrRet(0)

        .vspdData.Col = C_ItemName
        .vspdData.Text = arrRet(1)

        .vspdData.Col = C_Specification
        .vspdData.Text = arrRet(17)

        .vspdData.Col = C_SLCD
        .vspdData.Text = arrRet(10)

        .vspdData.Col = C_OrderUnitMFG
        .vspdData.Text = arrRet(9)
        .vspdData.Col = C_OrderUnit
        .vspdData.Text = arrRet(9)
        .vspdData.Col = C_BaseUnit
        .vspdData.Text = arrRet(2)
        frm1.txtOrderUnitMFG.value = arrRet(9)

        .vspdData.Col = C_OrderLtMFG
        .vspdData.Text = arrRet(3)
        frm1.txtOrderLtMFG.value = arrRet(3)

        .vspdData.Col = C_FixedMRPQty
        .vspdData.Text = arrRet(13)
        frm1.txtFixedMRPQty.value = arrRet(13)

        .vspdData.Col = C_ValidFromDT
        .vspdData.Text = arrRet(14)
        frm1.txtValidFromDT.Text = arrRet(14)

        .vspdData.Col = C_ValidToDT
        .vspdData.Text = arrRet(15)
        frm1.txtValidToDT.Text = arrRet(15)

        .vspdData.Col = C_MaxMrpQty
        .vspdData.Text = arrRet(5)
        frm1.txtMaxMRPQty.value = arrRet(5)

        .vspdData.Col = C_MinMrpQty
        .vspdData.Text = arrRet(4)
        frm1.txtMinMRPQty.value = arrRet(4)

        .vspdData.Col = C_RoundQty
        .vspdData.Text = arrRet(6)
        frm1.txtRoundQty.value = arrRet(6)

        .vspdData.Col = C_OrderLtMFG
        .vspdData.Text = arrRet(3)
        frm1.txtOrderLtMFG.value = arrRet(3)

        Call LookUpMajorRouting(arrRet(0), Row)

        ggoSpread.Source = .vspdData
        ggoSpread.UpdateRow .vspdData.ActiveRow

    End With

End Function

'------------------------------------------  SetMRPRunNo()  -------------------------------------------
'   Name : SetMRPRunNo()
'   Description : I
'---------------------------------------------------------------------------------------------------------
Function SetMRPRunNo(Byval arrRet)

    With frm1
        .txtMRPRunNo.value = arrRet(0)
    End With

End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'   Name : SetTrackingNo()
'   Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)

    frm1.txtTrackingNo.Value = arrRet(0)

End Function

'------------------------------------------  SetTrackingNo2()  -----------------------------------------
'   Name : SetTrackingNo2()
'   Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo2(Byval arrRet, Byval Row)

    With frm1

        .vspdData.Row = Row
        .vspdData.Col = C_TrackingNo
        .vspdData.Text = arrRet(0)
        Call vspdData_Change(.vspdData.Col, .vspdData.Row)

    End With

End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'   Name : SetConPlant()
'   Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
    frm1.txtPlantCd.Value    = arrRet(0)
    frm1.txtPlantNm.Value    = arrRet(1)

    Call LookUpInvClsDt()

End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'   Name : SetProdOrderNo()
'   Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value    = arrRet(0)
End Function

'------------------------------------------  SetRoutingNo()  --------------------------------------------------
'   Name : SetRoutingNo()
'   Description : RoutingNo Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetRoutingNo(Byval arrRet, Byval Row)

    With frm1
        .vspdData.Row = Row
        .vspdData.Col = C_RoutingNo
        .vspdData.Text = arrRet(0)
        .vspdData.Col = C_BOMNo
        .vspdData.Text = arrRet(2)
        Call vspdData_Change(.vspdData.Col, .vspdData.Row)

    End With

End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'   Name : SetSLCd()
'   Description : Ware House Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet, Byval Row)

    With frm1
        .vspdData.Row = Row
        .vspdData.Col = C_SLCD
        .vspdData.Text = arrRet(0)
        .vspdData.Col = C_SLNM
        .vspdData.Text = arrRet(1)
        Call vspdData_Change(.vspdData.Col, .vspdData.Row)
    End With

End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'   Name : SetUnit()
'   Description : Unit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetUnit(byval arrRet, Byval Row)

    With frm1
        .vspdData.Row = Row
        .vspdData.Col = C_OrderUnit
        .vspdData.Text = arrRet(0)
        .vspdData.Col = C_OrderUnitMfg
        .vspdData.Text = arrRet(0)
        Call vspdData_Change(.vspdData.Col, .vspdData.Row)
    End With

End Function

'------------------------------------------  SetCostCtr()  -----------------------------------------------
'   Name : SetCostCtr()
'   Description : Cost Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCostCtr(byval arrRet, byVal Row)

     With frm1
        .vspdData.Row = Row
        .vspdData.Col = C_CostCd
        .vspdData.Text = arrRet(0)
        .vspdData.Col = C_CostNm
        .vspdData.Text = arrRet(1)
        Call vspdData_Change(.vspdData.Col, .vspdData.Row)
    End With

End Function

'------------------------------------------  SetConWC2()  ----------------------------------------------
'	Name : SetConWC2()
'	Description : Work Center Popup for Grid 2 에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConWC2(Byval arrRet, Byval Row)
	With frm1
		.vspdData.Col = C_WcCd
		.vspdData.Text = UCase(arrRet(0))
		.vspdData.Col = C_WcNm
		.vspdData.Text = UCase(arrRet(1))
	End With
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'----------------------------------------  LookUpInvClsDt()  -------------------------------------------
'   Name : LookUpInvClsDt()
'   Description : LookUp Inventory Close Date
'---------------------------------------------------------------------------------------------------------

Function LookUpInvClsDt()

    Dim strVal

    If LayerShowHide(1) = False Then Exit Function

    strVal = BIZ_PGM_INC_CLS_DT & "?txtMode=" & parent.UID_M0001            '☜: 비지니스 처리 ASP의 상태
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)          '☆: 조회 조건 데이타

    Call RunMyBizASP(MyBizASP, strVal)                                      '☜: 비지니스 ASP 를 가동

End Function

'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'   Name : LookUpItem ByPlant()
'   Description : LookUp Item By Plant
'---------------------------------------------------------------------------------------------------------

Function LookUpItemByPlant(Byval strItemCd, Byval Row)

    Dim strSelect
    Dim gComNum1000, gComNumDec, gAPNum1000, gAPNumDec

    gComNum1000 = parent.gComNum1000
    gComNumDec = parent.gComNumDec
    gAPNum1000 = parent.gAPNum1000
    gAPNumDec = parent.gAPNumDec

    If strItemCd = "" Then Exit Function

    frm1.vspdData.Col = C_ItemCode
    frm1.vspdData.Row = Row

    strSelect = " a.ITEM_CD, a.BASIC_UNIT, a.ITEM_NM, a.SPEC, a.PHANTOM_FLG, a.VALID_FLG, b.PROCUR_TYPE, b.VALID_FLG, b.TRACKING_FLG, b.ORDER_UNIT_MFG, "
    strSelect = strSelect & "b.ORDER_LT_MFG, b.SCRAP_RATE_MFG, b.FIXED_MRP_QTY, b.MAX_MRP_QTY, b.MIN_MRP_QTY, b.ROUND_QTY, c.SL_CD, c.SL_NM, b.VALID_FROM_DT, b.VALID_TO_DT, "
    strSelect = strSelect & "dbo.ufn_GetCodeName(" & FilterVar("P1012", "''", "S") & ", b.MPS_MGR), dbo.ufn_GetCodeName(" & FilterVar("P1011", "''", "S") & ", b.MRP_MGR), dbo.ufn_GetCodeName(" & FilterVar("P1015", "''", "S") & ", b.PROD_MGR) "
    If  CommonQueryRs2by2(strSelect, " B_ITEM a, B_ITEM_BY_PLANT b, B_STORAGE_LOCATION c ", " a.ITEM_CD = b.ITEM_CD AND b.MAJOR_SL_CD *= c.SL_CD AND b.PLANT_CD = " & _
        FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND b.ITEM_CD = " & FilterVar(Frm1.vspdData.Text, "''", "S"), lgF0) = False Then
        Call DisplayMsgBox("122700","X", Frm1.vspdData.Text,"X")
        Call LookUpItemByPlantFail(Row)
        Exit Function
    End If

    lgF0 = Split(lgF0, Chr(11))

    With frm1.vspdData

        If lgF0(6) = "N" Then 'Invalid Item
            Call DisplayMsgBox("122623", "X", "X", "X")
            Call LookUpItemByPlantFail(Row)

        ElseIf lgF0(8) = "N" Then 'Invalid Item by plant
            Call DisplayMsgBox("122729", "X", "X", "X")
            Call LookUpItemByPlantFail(Row)

        Else
            If lgF0(9) = "N" Then 'Tracking Flg
                ggoSpread.SpreadLock C_TrackingNo, Row, C_TrackingNoPopup, Row
                ggoSpread.SSSetProtected C_TrackingNo,  Row, Row
                ggoSpread.SSSetProtected C_TrackingNoPopup, Row, Row
                Call .SetText(C_TrackingNo, Row, "*")
            Else
                ggoSpread.SpreadUnLock C_TrackingNo, Row, C_TrackingNoPopup, Row
                ggoSpread.SSSetRequired C_TrackingNo, Row, Row
                Call .SetText(C_TrackingNo, Row, "")
            End If

            If lgF0(5) = "Y" Then ' Phantom
                Call DisplayMsgBox("189214", "X", "X", "X")
                Call LookUpItemByPlantFail(Row)
            Else
                Call .SetText(C_ItemName, Row, lgF0(3))
                Call .SetText(C_Specification, Row, lgF0(4))
                Call .SetText(C_OrderUnit, Row, lgF0(10))
                Call .SetText(C_BaseUnit, Row, lgF0(2))
                Call .SetText(C_SLCD, Row, lgF0(17))
                Call .SetText(C_SLNM, Row, lgF0(18))
                Call .SetText(C_ValidFromDT, Row, UNIConvDateDBToCompany(lgF0(19), ""))
                Call .SetText(C_ValidToDT, Row, UNIConvDateDBToCompany(lgF0(20), ""))
                Call .SetText(C_MPSMgr, Row, lgF0(21))
                Call .SetText(C_MRPMgr, Row, lgF0(22))
                Call .SetText(C_ProdMgr, Row, lgF0(23))

                Call .SetText(C_OrderLtMFG, Row, uniConvNumAToB(lgF0(11), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X"))

                Call .SetText(C_ScrapRateMFG, Row, uniConvNumAToB(lgF0(12), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X"))
                Call .SetText(C_FixedMRPQty, Row, uniConvNumAToB(lgF0(13), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X"))
                Call .SetText(C_MaxMRPQty, Row, uniConvNumAToB(lgF0(14), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X"))
                Call .SetText(C_MinMRPQty, Row, uniConvNumAToB(lgF0(15), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X"))
                Call .SetText(C_RoundQty, Row, uniConvNumAToB(lgF0(16), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X"))

                frm1.txtOrderUnitMFG.value  = lgF0(10)
                frm1.txtOrderLtMFG.Text     = uniConvNumAToB(lgF0(11), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X")
                frm1.txtScrapRateMfg.Text   = uniConvNumAToB(lgF0(12), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X")
                frm1.txtFixedMRPQty.Text    = uniConvNumAToB(lgF0(13), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X")
                frm1.txtMaxMRPQty.Text      = uniConvNumAToB(lgF0(14), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X")
                frm1.txtMinMRPQty.Text      = uniConvNumAToB(lgF0(15), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X")
                frm1.txtRoundQty.Text       = uniConvNumAToB(lgF0(16), gAPNum1000, gAPNumDec, gComNum1000, gComNumDec, True, "X", "X")
                frm1.txtValidFromDT.Text    = UNIConvDateDBToCompany(lgF0(19), "")
                frm1.txtValidToDT.Text      = UNIConvDateDBToCompany(lgF0(20), "")
                frm1.txtMPSMgr.value        = lgF0(21)
                frm1.txtMRPMgr.value        = lgF0(22)
                frm1.txtProdMgr.value       = lgF0(23)

                Call LookUpMajorRouting(strItemCd, Row)
            End If
        End If

    End With

End Function

Function LookUpItemByPlantFail(Byval Row)

    With frm1.vspdData
        Call .SetText(C_ItemCode, .Row, "")
        Call .SetText(C_ItemName, .Row, "")
        Call .SetText(C_Specification, .Row, "")
        Call .SetText(C_OrderUnit, .Row, "")
        Call .SetFloat(C_OrderLtMFG, .Row, 0)
        Call .SetFloat(C_ScrapRateMFG, .Row, 0)
        Call .SetFloat(C_MaxMRPQty, .Row, 0)
        Call .SetFloat(C_MinMRPQty, .Row, 0)
        Call .SetFloat(C_RoundQty, .Row, 0)
        Call .SetText(C_SLCD, .Row, "")
        Call .SetText(C_SLNM, .Row, "")
        Call .SetText(C_BaseUnit, .Row, "")
    End With

End Function

'-------------------------------------  LookUpItemByPlant Success()  ---------------------------------------
'   Name : LookUpItemByPlantSuccess()
'   Description : LookUp Item By Plant Success
'---------------------------------------------------------------------------------------------------------
Function LookUpItemByPlantSuccess(Byval strItemCd, Byval Row)
    Call LookUpMajorRouting(strItemCd, Row)
End Function

'-------------------------------------  LookUpMajorRouting()  -----------------------------------------
'   Name : LookUpMajorRouting()
'   Description : LookUp Major Routing
'---------------------------------------------------------------------------------------------------------
Function LookUpMajorRouting(Byval strItemCd, Byval Row)

    If  CommonQueryRs("A.ROUT_NO, A.COST_CD, B.COST_NM ", "P_ROUTING_HEADER A , B_COST_CENTER B ", _
                " A.PLANT_CD *= B.PLANT_CD AND A.COST_CD *= B.COST_CD AND  A.MAJOR_FLG = " & FilterVar("Y", "''", "S") & "  AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(strItemCd, "''", "S"), _
                lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then
        Exit Function
    End If

    lgF0 = Split(lgF0, Chr(11))
    lgF1 = Split(lgF1, Chr(11))
    lgF2 = Split(lGF2, Chr(11))

    Call frm1.vspdData.SetText(C_RoutingNo, Row, lgF0(0))
    Call frm1.vspdData.SetText(C_CostCd, Row, lgF1(0))
    Call frm1.vspdData.SetText(C_CostNm, Row, lgF2(0))

    Call LookUpWcCd(strItemCd, lgF0(0), Row)

End Function

'-------------------------------------  LookUpRouting()  -----------------------------------------
'   Name : LookUpRouting()
'   Description : LookUp Major Routing
'---------------------------------------------------------------------------------------------------------
Function LookUpRouting(Byval strItemCd, Byval strRouting, Byval Row)

    If  CommonQueryRs("A.ROUT_NO, A.COST_CD, B.COST_NM ", "P_ROUTING_HEADER A , B_COST_CENTER B ", _
                " A.PLANT_CD *= B.PLANT_CD AND A.COST_CD *= B.COST_CD AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & _
                " AND A.ITEM_CD = " & FilterVar(strItemCd, "''", "S") & " AND A.ROUT_NO = " & FilterVar(strRouting, "''", "S") , _
                lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then
        Exit Function
    End If

    lgF0 = Split(lgF0, Chr(11))
    lgF1 = Split(lgF1, Chr(11))
    lgF2 = Split(lgF2, Chr(11))

    Call frm1.vspdData.SetText(C_RoutingNo, Row, lgF0(0))
    Call frm1.vspdData.SetText(C_CostCd, Row, lgF1(0))
    Call frm1.vspdData.SetText(C_CostNm, Row, lgF2(0))

    Call LookUpWcCd(strItemCd, strRouting, Row)
End Function

'-------------------------------------  LookUpWcCd()  -----------------------------------------
'   Name : LookUpWcCd()
'   Description : LookUp Wc Cd And Opr No
'---------------------------------------------------------------------------------------------------------
Function LookUpWcCd(Byval strItemCd, Byval strRouting, Byval Row)

    If  CommonQueryRs("A.ROUT_NO, A.OPR_NO, A.WC_CD, B.WC_NM ", "P_ROUTING_DETAIL A, P_WORK_CENTER B ", _
                " A.PLANT_CD = B.PLANT_CD AND A.WC_CD = B.WC_CD AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & _
                " AND A.ITEM_CD = " & FilterVar(strItemCd, "''", "S") & " AND A.ROUT_NO = " & FilterVar(strRouting, "''", "S") , _
                lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then
        Exit Function
    End If

'    lgF0 = Split(lgF0, Chr(11))
    lgF1 = Split(lgF1, Chr(11))
    lgF2 = Split(lgF2, Chr(11))
    lgF3 = Split(lgF3, Chr(11))

    Call frm1.vspdData.SetText(C_OprNo, Row, lgF1(0))
    Call frm1.vspdData.SetText(C_WcCd, Row, lgF2(0))
    Call frm1.vspdData.SetText(C_WcNm, Row, lgF3(0))

End Function

'-------------------------------------  LookUpWcCd()  -----------------------------------------
'   Name : LookUpWcCd()
'   Description : LookUp Wc Cd And Opr No
'---------------------------------------------------------------------------------------------------------
Function TabWcCd(Byval strWcCd, Byval Row)

    If  CommonQueryRs("WC_CD, WC_NM ", "P_WORK_CENTER ", _
                " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & _
                " AND WC_CD = " & FilterVar(strWcCd, "''", "S") , _
                lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then

        Call DisplayMsgBox("182100", "x", "x", "x")
        Call frm1.vspdData.SetText(C_WcCd, Row, "")
        Call frm1.vspdData.SetText(C_WcNm, Row, "")
        Exit Function
    End If

    lgF0 = Split(lgF0, Chr(11))
    lgF1 = Split(lgF1, Chr(11))

    Call frm1.vspdData.SetText(C_WcCd, Row, lgF0(0))
    Call frm1.vspdData.SetText(C_WcNm, Row, lgF1(0))


End Function

'-------------------------------------  LookUpWcCd()  -----------------------------------------
'   Name : LookUpWcCd()
'   Description : LookUp Wc Cd And Opr No
'---------------------------------------------------------------------------------------------------------
Function TabSLCd(Byval strSlCd, ByVal flag, Byval Row)

    If  CommonQueryRs("SL_CD, SL_NM ", "B_STORAGE_LOCATION ", _
                " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & _
                " AND SL_CD = " & FilterVar(strSlCd, "''", "S") , _
                lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then

        Call DisplayMsgBox("169931", "x", "x", "x")
        If flag = "TOP" Then      
            Call frm1.vspdData.SetText(C_SLCD, Row, "")
            Call frm1.vspdData.SetText(C_SLNM, Row, "")
        Else
            Call frm1.vspdData2.SetText(C_SlCd2, Row, "")
            Call frm1.vspdData2.SetText(C_SlNm2, Row, "")
        End If
        Exit Function
    End If

    lgF0 = Split(lgF0, Chr(11))
    lgF1 = Split(lgF1, Chr(11))

    If flag = "TOP" Then      
        Call frm1.vspdData.SetText(C_SLCD, Row, lgF0(0))
        Call frm1.vspdData.SetText(C_SLNM, Row, lgF1(0))
    Else
        ggoSpread.Source = frm1.vspdData2
        Call frm1.vspdData2.SetText(C_SlCd2, Row, lgF0(0))
        Call frm1.vspdData2.SetText(C_SlNm2, Row, lgF1(0))
        ggoSpread.UpdateRow Row        
    End If



End Function

'-------------------------------------  LookUpDate()  -----------------------------------------
'   Name : LookUpDate()
'   Description : LookUp Major Routing
'---------------------------------------------------------------------------------------------------------
Function LookUpDate(Byval strType, Byval LngProdLt, Byval DtPlanDt, Byval Row)

    If strType = "START_DATE" Then
        LngProdLt = 0 - CInt(Trim(LngProdLt))
    Else
        LngProdLt = Trim(LngProdLt)
    End If

    If LngProdLt = 0 Then
        If strType = "START_DATE" Then
            lgPlannedDate = DtPlanDt
        Else
            lgPlannedDate = DtPlanDt
        End If
        Call LookUpDateSuccess(strType, Row)
        Exit Function
    End If

    If CommonQueryRs("a.DT", "P_MFG_CALENDAR a", "a.CAL_TYPE = " & FilterVar(lgCalType, "''", "S") & _
       " AND a.TOT_ACCUM_WORK_DAY = (SELECT b.TOT_ACCUM_WORK_DAY - " & LngProdLt & " FROM P_MFG_CALENDAR b WHERE b.CAL_TYPE = " & FilterVar(lgCalType, "''", "S") & _
       " AND b.DT =  " & FilterVar(UniConvDate(DtPlanDt), "''", "S") & ") AND a.WORK_TYPE <> 0", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then
        Exit Function
    End If

    lgF0 = Split(lgF0, Chr(11))
    lgPlannedDate = UNIDateClientFormat(lgF0(0))
    Call LookUpDateSuccess(strType, Row)

End Function

'-------------------------------------  LookUpDateSuccess()  -----------------------------------------
'   Name : LookUpDateSuccess()
'   Description : LookUp Major Routing
'---------------------------------------------------------------------------------------------------------
Function LookUpDateSuccess(Byval strType, Byval Row)

    frm1.vspdData.Row = Row
    If strType = "START_DATE" Then
        frm1.vspdData.Col = C_PlanEndDt
        frm1.vspdData.Text = lgPlannedDate
    Else
        If UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat) <> "" and lgPlannedDate <> "" Then
            If UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat) >= UniConvDateAToB(lgPlannedDate, parent.gDateFormat, parent.gServerDateFormat) Then
                Call LookUpDate("END_DATE", -1, lgInvCloseDt, Row)
            End If
        End If
        frm1.vspdData.Col = C_PlanEndDt
        If UniConvDateAToB(frm1.vspdData.Text, parent.gDateFormat, parent.gServerDateFormat) < UniConvDateAToB(lgPlannedDate, parent.gDateFormat, parent.gServerDateFormat) Then
            frm1.vspdData.Col = C_PlanStartDt
            frm1.vspdData.Text = frm1.vspdData.Text
        Else
            frm1.vspdData.Col = C_PlanStartDt
            frm1.vspdData.Text = lgPlannedDate
        End If

    End If

End Function

'Add 2005-09-27
Sub ProtectCostCd()
    If UCase(Trim(Frm1.hOprCostFlag.value)) = "Y" Then
        Call InitTrackingNCost("C")
    Else
        ggoSpread.Source = frm1.vspdData
        ggoSpread.SpreadLock C_CostCd, -1, C_CostPopUp
        ggoSpread.SSSetProtected C_CostCd,  -1
        ggoSpread.SSSetProtected C_CostPopUp,  -1
    End If
End Sub

'-------------------------------------  ReleaseOrder()  -----------------------------------------
'   Name : ReleaseOrder()
'   Description : Release Order
'---------------------------------------------------------------------------------------------------------
Function ReleaseOrder()

    Dim IntRetCD
    Dim strVal
    Dim IntRows

    Dim iColSep, iRowSep

    Dim strCUTotalvalLen                    '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규]

    Dim iFormLimitByte                      '102399byte

    Dim objTEXTAREA                         '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼

    Dim iTmpCUBuffer                        '현재의 버퍼 [수정,신규]
    Dim iTmpCUBufferCount                   '현재의 버퍼 Position
    Dim iTmpCUBufferMaxCount                '현재의 버퍼 Chunk Size

    If lgIntFlgMode = parent.OPMD_CMODE Then
        Call DisplayMsgBox("900002", "x", "x", "x")
        Exit Function
    End If

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    Dim firstCheck
    Dim secondCheck
    firstCheck = "N"
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        firstCheck = "Y"
    End If

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        secondCheck ="Y"
    End If

    If firstCheck = "Y" Or secondCheck = "Y" Then
        IntRetCD = DisplayMsgBox("189217", "x", "x", "x")   '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")

    If IntRetCD = vbNo Then
        Exit Function
    End If

    If LayerShowHide(1) = False Then Exit Function

    ggoSpread.Source = frm1.vspdData

    iColSep = parent.gColSep : iRowSep = parent.gRowSep

    '한번에 설정한 버퍼의 크기 설정
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT

    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE

    '버퍼의 초기화
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)

    iTmpCUBufferCount = -1

    strCUTotalvalLen = 0


    For IntRows = ggoSpread.Source.SelBlockRow To ggoSpread.Source.SelBlockRow2

        strVal = ""

        strVal = strVal & "CREATE" & iColSep
        strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
        frm1.vspdData.Row = IntRows
        frm1.vspdData.Col = C_ProdtOrderNo
        strVal = strVal & UCase(Trim(frm1.vspdData.Text)) & iColSep
        strVal = strVal & IntRows & iRowSep

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

    Next

    If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리
       Set objTEXTAREA = document.createElement("TEXTAREA")
       objTEXTAREA.name   = "txtCUSpread"
       objTEXTAREA.value = Join(iTmpCUBuffer,"")
       divTextArea.appendChild(objTEXTAREA)
    End If

    Call ExecMyBizASP(frm1, BIZ_PGM_RELEASE_ID)                                     '☜: 비지니스 ASP 를 가동

End Function

'-------------------------------------  CancelOrder()  -----------------------------------------
'   Name : CancelOrder()
'   Description : Cancel Order
'---------------------------------------------------------------------------------------------------------
Function CancelOrder() 
    Dim IntRetCD
    Dim strVal
    Dim IntRows

    Dim iColSep, iRowSep

    Dim strCUTotalvalLen                    '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규]

    Dim iFormLimitByte                      '102399byte

    Dim objTEXTAREA                         '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼

    Dim iTmpCUBuffer                        '현재의 버퍼 [수정,신규]
    Dim iTmpCUBufferCount                   '현재의 버퍼 Position
    Dim iTmpCUBufferMaxCount                '현재의 버퍼 Chunk Size

    If lgIntFlgMode = parent.OPMD_CMODE Then
        Call DisplayMsgBox("900002", "x", "x", "x")
        Exit Function
    End If

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    Dim firstCheck
    Dim secondCheck
    firstCheck = "N"
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        firstCheck = "Y"
    End If

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        secondCheck ="Y"
    End If

    If firstCheck = "Y" Or secondCheck = "Y" Then
        IntRetCD = DisplayMsgBox("189217", "x", "x", "x")   '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")

    If IntRetCD = vbNo Then
        Exit Function
    End If

    If LayerShowHide(1) = False Then Exit Function

    ggoSpread.Source = frm1.vspdData

    iColSep = parent.gColSep : iRowSep = parent.gRowSep

    '한번에 설정한 버퍼의 크기 설정
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT

    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE

    '버퍼의 초기화
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)

    iTmpCUBufferCount = -1

    strCUTotalvalLen = 0


    For IntRows = ggoSpread.Source.SelBlockRow To ggoSpread.Source.SelBlockRow2
        strVal = ""
        strVal = strVal & UCase(Trim(frm1.hPlantCd.value)) & iColSep			

        frm1.vspdData.Row = IntRows
        frm1.vspdData.Col = C_ProdtOrderNo			
        strVal = strVal & Trim(frm1.vspdData.Text) & iColSep
        strVal = strVal & lRow & iRowSep  

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

    Next

    If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리
       Set objTEXTAREA = document.createElement("TEXTAREA")
       objTEXTAREA.name   = "txtCUSpread"
       objTEXTAREA.value = Join(iTmpCUBuffer,"")
       divTextArea.appendChild(objTEXTAREA)
    End If

    Call ExecMyBizASP(frm1, BIZ_PGM_CANCEL_ID)                                     '☜: 비지니스 ASP 를 가동

End Function



Function JumpOrderRun()

    Dim IntRetCd, strVal

    If lgIntFlgMode = parent.OPMD_CMODE Then
        Call DisplayMsgBox("900002", "x", "x", "x")
        Exit Function
    End If

    Dim firstCheck
    Dim secondCheck
    firstCheck = "N"
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        firstCheck = "Y"
    End If

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        secondCheck ="Y"
    End If

    If firstCheck = "Y" Or secondCheck = "Y" Then
        IntRetCD = DisplayMsgBox("189217", "x", "x", "x")   '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    If frm1.txtPlantCd.value= "" Then
        Call DisplayMsgBox("971012","X", "공장","X")
        frm1.txtPlantCd.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If

    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_ReWorkFlag
    If frm1.vspdData.Text = "Y" Then
        Call DisplayMsgBox("189218", "x", "x", "x")
        Exit Function
    End If

    WriteCookie "txtPlantCd", UCase(Trim(frm1.txtPlantCd.value))
    WriteCookie "txtPlantNm", frm1.txtPlantNm.value
    frm1.vspdData.Col = C_ItemCode
    WriteCookie "txtItemCd", UCase(Trim(frm1.vspdData.Text))
    frm1.vspdData.Col = C_ItemName
    WriteCookie "txtItemNm", Trim(frm1.vspdData.Text)
    frm1.vspdData.Col = C_Specification
    WriteCookie "txtSpecification", Trim(frm1.vspdData.Text)
    frm1.vspdData.Col = C_ProdtOrderNo
    WriteCookie "txtProdOrderNo", UCase(Trim(frm1.vspdData.Text))
    frm1.vspdData.Col = C_PlanOrderNo
    WriteCookie "txtPlanOrderNo", UCase(Trim(frm1.vspdData.Text))
    frm1.vspdData.Col = C_OrderQty
    WriteCookie "txtOrderQty", UCase(Trim(frm1.vspdData.Text))
    frm1.vspdData.Col = C_OrderUnit
    WriteCookie "txtOrderUnit", UCase(Trim(frm1.vspdData.Text))
    frm1.vspdData.Col = C_PlanStartDt
    WriteCookie "txtPlanStartDt", UCase(Trim(frm1.vspdData.Text))
    frm1.vspdData.Col = C_PlanEndDt
    WriteCookie "txtPlanEndDt", UCase(Trim(frm1.vspdData.Text))
    WriteCookie "txtInvCloseDt", lgInvCloseDt
    WriteCookie "txtPGMID", "P4112MA1"

    navigate BIZ_PGM_JUMPORDERRUN_ID

End Function


'#########################################################################################################
'                                               3. Event부
'   기능: Event 함수에 관한 처리
'   설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'   Window에 발생 하는 모든 Even 처리
'*********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtPlantCd_onChange()
'   Event Desc :
'=======================================================================================================
Sub txtPlantCd_onChange()
     Dim IntRetCd

    If  frm1.txtPlantCd.value = "" Then
        frm1.txtPlantCd.Value = ""
        frm1.txtPlantNm.Value = ""
        frm1.hOprCostFlag.value = ""
    Else

        Call LookUpInvClsDt()

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

        Call ProtectCostCd()

     End If
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'   Document의 TAG에서 발생 하는 Event 처리
'   Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나
'   Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag 처리  *********************************************
'   Window에 발생 하는 모든 Even 처리
'*********************************************************************************************************

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim DtPlanStartDt, DtPlanComptDt, DtInvCloseDt
    Dim strYear,strMonth,strDay
    Dim DtPlanStartDtDateFormat, DtPlanComptDtDateFormat
    Dim strItemCd

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    With frm1.vspdData

    Select Case Col

        Case C_ItemCode

            frm1.vspdData.Col = C_ItemCode
            Call LookUpItemByPlant(.Value, Row)
        Case C_WcCd

            frm1.vspdData.Col = C_WcCd
            Call TabWcCd(.Value, Row)
        Case C_SLCD

            frm1.vspdData.Col = C_SLCD
            Call TabSLCd(.Value, "TOP", Row)
        Case C_OrderQty

            .Col = C_OrderQty
            If .Value = 0 Then
                Call DisplayMsgBox("189208", "x", "x", "x")
                .Value = ""
                .Focus
                Set gActiveElement = document.activeElement
                Exit Sub
            End If

            If .Value < 0 Then
                Call DisplayMsgBox("189208", "x", "x", "x")
                .Value = ""
                Exit Sub
            End If

            .Col = C_OrderQtyInBaseUnit
            .Value = 0

        Case C_OrderUnit

            .Col = C_OrderQtyInBaseUnit
            .Value = 0

        Case C_PlanStartDt

            DtInvCloseDt = UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat)

            .Col = C_PlanEndDt
            DtPlanComptDt = UniConvDateAToB(.Text, parent.gDateFormat, parent.gServerDateFormat)
            .Col = C_PlanStartDt
            DtPlanStartDt = UniConvDateAToB(.Text, parent.gDateFormat, parent.gServerDateFormat)
            DtPlanStartDtDateFormat = .Text
            If (DtPlanStartDt <> "" and DtPlanComptDt <> "") and (isdate(DtPlanStartDt) and isdate(DtPlanComptDt)) Then
                If DtPlanStartDt > DtPlanComptDt  Then
                    Call DisplayMsgBox("189207", "x", "x", "x")
                    .Col = C_PlanStartDt
                    .Text = ""
                    Exit Sub
                End If
            End If

            If (DtPlanStartDt <> "" and DtInvCloseDt <> "") and (isdate(DtPlanStartDt) and isdate(DtInvCloseDt)) Then
                If DtPlanStartDt <= DtInvCloseDt Then
                    Call DisplayMsgBox("189204", "x", "x", "x")
                    .Col = C_PlanStartDt
                    .Text = ""
                    Exit Sub
                End If
            End If

            .Col = C_PlanEndDt
            If .Text = "" Then
                .Col = C_OrderLtMFG
                Call LookUpDate("START_DATE", .Text, DtPlanStartDtDateFormat, Row)
            End If
        Case C_PlanEndDt

            DtInvCloseDt = UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat)

            .Col = C_PlanStartDt
            DtPlanStartDt = UniConvDateAToB(.Text, parent.gDateFormat, parent.gServerDateFormat)
            .Col = C_PlanEndDt
            DtPlanComptDt = UniConvDateAToB(.Text, parent.gDateFormat, parent.gServerDateFormat)
            DtPlanComptDtDateFormat = .Text

            If (DtPlanComptDt <> "" and DtPlanStartDt <> "") and (isdate(DtPlanComptDt) and isdate(DtPlanStartDt)) Then
                If DtPlanStartDt > DtPlanComptDt Then
                    Call DisplayMsgBox("189207", "x", "x", "x")
                    .Col = C_PlanEndDt
                    .Text = ""
                    Exit Sub
                End If
            End If

            If (DtPlanComptDt <> "" and DtInvCloseDt <> "") and (isdate(DtPlanComptDt) and isdate(DtInvCloseDt)) Then
                If DtPlanComptDt <= DtInvCloseDt Then
                    Call DisplayMsgBox("189205", "x", "x", "x")
                    .Col = C_PlanEndDt
                    .Text = ""
                    Exit Sub
                End If
            End If

            .Col = C_PlanStartDt
            If .Text = "" Then
                .Col = C_OrderLtMFG
                Call LookUpDate("END_DATE", .Text, DtPlanComptDtDateFormat, Row)
            End If

        'Add 2004-10-04
        Case C_RoutingNo
            frm1.vspdData.Col = C_ItemCode
            strItemCd = .Value
            frm1.vspdData.Col = C_RoutingNo
            Call LookUpRouting(strItemCd, .Value, Row)

    End Select

    End With

End Sub


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    Dim secondCheck
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer

    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        secondCheck ="Y"
    End If

    If secondCheck = "Y" Then
        IntRetCD = DisplayMsgBox("189217", "x", "x", "x")   '⊙: Display Message(There is no changed data.)
        Exit Sub
    End If

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer

    whereVspdData = "TOP"

    If lgIntFlgMode = Parent.OPMD_UMODE Then
        Call SetPopupMenuItemInf("1101111111")         '화면별 설정
    Else
        Call SetPopupMenuItemInf("1001111111")         '화면별 설정
    End If

    With frm1.vspdData
        '----------------------
        'Column Split
        '----------------------
        gMouseClickStatus = "SPC"

        Set gActiveSpdSheet = frm1.vspdData

        If frm1.vspdData.MaxRows = 0 Then
            Exit Sub
        End If

        If Row <= 0 Then
            ggoSpread.Source = frm1.vspdData
            If lgSortKey = 1 Then
                ggoSpread.SSSort Col                    'Sort in Ascending
                lgSortKey = 2
            Else
                ggoSpread.SSSort Col, lgSortKey     'Sort in Descending
                lgSortKey = 1
            End If

            lgOldRow = Row

            frm1.vspdData2.MaxRows = 0

            If DbDtlQuery = False Then
                Call RestoreToolBar()
                Exit Sub
            End If
        Else
            If lgOldRow <> Row Then

                frm1.vspdData.Col = 1
                frm1.vspdData.Row = row

                lgOldRow = Row

                frm1.vspdData2.MaxRows = 0

                If DbDtlQuery = False Then

                    Call RestoreToolBar()
                    Exit Sub
                End If

            End If
        End If

        '------ Developer Coding part (Start)
        .Row = .ActiveRow
        .Col = C_OrderUnitMFG
        frm1.txtOrderUnitMFG.value = .Text
        .Col = C_MinMRPQty
        frm1.txtMinMRPQty.value = .Text
        .Col = C_FixedMRPQty
        frm1.txtFixedMRPQty.value = .Text
        .Col = C_MaxMRPQty
        frm1.txtMaxMRPQty.value = .Text
        .Col = C_RoundQty
        frm1.txtRoundQty.value = .Text
        .Col = C_ValidFromDT
        frm1.txtValidFromDT.Text = .Text
        .Col = C_ValidToDT
        frm1.txtValidToDT.Text = .Text
        .Col = C_OrderLtMFG
        frm1.txtOrderLtMFG.value = .Text
        .Col = C_ScrapRateMFG
        frm1.txtScrapRateMFG.value = .Text
        .Col = C_MPSMgr
        frm1.txtMPSMgr.value = .Text
        .Col = C_MRPMgr
        frm1.txtMRPMgr.value = .Text
        .Col = C_ProdMgr
        frm1.txtProdMgr.value = .Text
        '------ Developer Coding part (End)

    End With

End Sub


Sub vspdData2_Click(ByVal Col , ByVal Row )
    whereVspdData = "DOWN"

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	End If
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2
    
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
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
	
 	End If

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
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName

    If Row <= 0 Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
    '------ Developer Coding part (Start)
    '------ Developer Coding part (End)
End Sub

'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================

Sub vspddata_KeyPress(index , KeyAscii )

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
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex

     '----------  Coding part  -------------------------------------------------------------
     ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직
    With frm1.vspdData

        .Row = Row

        Select Case Col
            Case  1
                .Col = Col
                intIndex = .Value
                .Col = C_BillFG
                .Value = intIndex
        End Select
    End With
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If CheckRunningBizProcess = True Then           '⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음
         Exit Sub
    End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

     '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevKey <> "" Then                          '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음
            If DbQuery = False Then
                Call RestoreToolBar()
                Exit Sub
            End If
        End If
    End if

End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
        If Row < 1 Then Exit Sub

        Select Case Col

            Case C_ItemPopup
                .Col = C_ItemCode
                .Row = Row
                Call OpenItemInfo2(.Text, Row)
                Call SetActiveCell(frm1.vspdData,C_ItemCode,Row,"M","X","X")
                Set gActiveElement = document.activeElement

            Case C_TrackingNoPopup
                .Col = C_TrackingNo
                .Row = Row
                Call OpenTrackingInfo2(.Text, Row)
                Call SetActiveCell(frm1.vspdData,C_TrackingNo,Row,"M","X","X")
                Set gActiveElement = document.activeElement

            Case C_RoutingNoPopup
                .Col = C_RoutingNo
                .Row = Row
                Call OpenRoutingNo(.Text, Row)
                Call SetActiveCell(frm1.vspdData,C_RoutingNo,Row,"M","X","X")
                Set gActiveElement = document.activeElement

            Case C_SLCDPopup
                .Col = C_SLCD
                .Row = Row
                Call OpenSLCD(.Text, Row)
                Call SetActiveCell(frm1.vspdData,C_SLCD,Row,"M","X","X")
                Set gActiveElement = document.activeElement

            Case C_OrderUnitPopup
                .Col = C_OrderUnit
                .Row = Row
                Call OpenUnit(.Text, Row)
                Call SetActiveCell(frm1.vspdData,C_OrderUnit,Row,"M","X","X")
                Set gActiveElement = document.activeElement

            Case C_CostPopup
                .Col = C_CostCd
                .Row = Row
                Call OpenCostCtr(.Text, Row)
                Call SetActiveCell(frm1.vspdData,C_CostCd,Row,"M","X","X")
                Set gActiveElement = document.activeElement

            Case C_WcCdPopup
                .Col = C_WcCd
                .Row = Row
                Call OpenConWC2(.Text, Row)
                Call SetActiveCell(frm1.vspdData,C_WcCd,Row,"M","X","X")
                Set gActiveElement = document.activeElement
        End Select

    End With

End Sub


'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    Call InitSpreadComboBox
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.ReOrderingSpreadData
    If gActiveSpdSheet.Id = "A" Then Call InitData(1)

    If gActiveSpdSheet.Id = "A" Then
        With frm1.vspdData
            .ReDraw = False
            ggoSpread.SSSetProtected C_ProdtOrderNo, -1, -1
            ggoSpread.SSSetProtected C_ItemCode, -1, -1
            ggoSpread.SSSetProtected C_ItemPopup, -1, -1
            ggoSpread.SSSetProtected C_ProdtOrderNo, -1, -1
            ggoSpread.SSSetProtected C_TrackingNo, -1, -1
            ggoSpread.SSSetProtected C_CostCd, -1, -1
        
            If frm1.rdoStatus2.checked = True Then
                ggoSpread.SSSetProtected C_ReWorkFlag, -1, -1
                ggoSpread.SSSetProtected  C_CostPopup, -1, -1
                ggoSpread.SpreadUnLock  C_WcCd, -1, -1
                ggoSpread.SpreadUnLock  C_Remark, -1, -1
'20080227::hanc                ggoSpread.SSSetRequired C_WcCd, -1, -1

                For LngRow = 1 To .MaxRows
                    .Row = LngRow
                    .Col = C_TrackingNo
                    If .Text = "*" Or .Text = "" Then
                        ggoSpread.SpreadLock C_TrackingNo, LngRow, C_TrackingNoPopup, LngRow
                        ggoSpread.SSSetProtected C_TrackingNo, LngRow, LngRow
                        ggoSpread.SSSetProtected C_TrackingNoPopup, LngRow, LngRow
                    Else
                        ggoSpread.SpreadUnLock C_TrackingNo, LngRow, C_TrackingNoPopup, LngRow
                        ggoSpread.SSSetRequired C_TrackingNo, LngRow, LngRow
                    End If

                Next
            Else
                ggoSpread.SSSetProtected C_OrderQty, -1, -1
                ggoSpread.SSSetProtected C_BATCHCNT, -1, -1     '20080307
                ggoSpread.SSSetProtected C_OrderUnit, -1, -1
                ggoSpread.SSSetProtected C_OrderUnitPopup, -1, -1
                ggoSpread.SSSetProtected C_PlanStartDt, -1, -1
                ggoSpread.SSSetProtected C_PlanEndDt, -1, -1
                ggoSpread.SSSetProtected C_RoutingNo, -1, -1
                ggoSpread.SSSetProtected C_RoutingNoPopup, -1, -1
                ggoSpread.SSSetProtected C_SLCD, -1, -1
                ggoSpread.SSSetProtected C_SLCDPopup, -1, -1
                ggoSpread.SSSetProtected C_ReWorkFlag, -1, -1
                ggoSpread.SSSetProtected  C_CostPopup, -1, -1
                ggoSpread.SSSetProtected  C_WcCdPopup, 1, -1
                ggoSpread.SSSetProtected  C_TrackingNoPopup, -1, -1
            End If

            If .MaxRows < 1 Then Exit Sub

            If lgIntFlgMode = parent.OPMD_CMODE Then

                .Row = 1
                .Col = C_OrderUnitMFG
                frm1.txtOrderUnitMFG.value = .Text
                .Col = C_MinMRPQty
                frm1.txtMinMRPQty.value = .Text
                .Col = C_FixedMRPQty
                frm1.txtFixedMRPQty.value = .Text
                .Col = C_MaxMRPQty
                frm1.txtMaxMRPQty.value = .Text
                .Col = C_RoundQty
                frm1.txtRoundQty.value = .Text
                .Col = C_ValidFromDT
                frm1.txtValidFromDT.Text = .Text
                .Col = C_ValidToDT
                frm1.txtValidToDT.Text = .Text
                .Col = C_OrderLtMFG
                frm1.txtOrderLtMFG.value = .Text
                .Col = C_ScrapRateMFG
                frm1.txtScrapRateMFG.value = .Text
                .Col = C_MPSMgr
                frm1.txtMPSMgr.value = .Text
                .Col = C_MRPMgr
                frm1.txtMRPMgr.value = .Text
                .Col = C_ProdMgr
                frm1.txtProdMgr.value = .Text
            End If
            .ReDraw = True
        End With
    End If
    
    If gActiveSpdSheet.Id = "B" Then
        ggoSpread.SSSetProtected C_ItemCd2, -1, -1
        ggoSpread.SpreadUnLock C_SlCdPopup2, -1, -1
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
        SetFocusToDocument("M")
        Frm1.txtProdFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7
        SetFocusToDocument("M")
        Frm1.txtProdToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtProdFromDt_KeyDown(keycode, shift)
    If keycode = 13 Then
        Call MainQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtProdToDt_KeyDown(keycode, shift)
    If keycode = 13 Then
        Call MainQuery()
    End If
End Sub


'#########################################################################################################
'                                               4. Common Function부
'   기능: Common Function
'   설명: 환율처리함수, VAT 처리 함수
'#########################################################################################################


'#########################################################################################################
'                                               5. Interface부
'   기능: Interface
'   설명: 각각의 Toolbar에 대한 처리를 행한다.
'         Toolbar의 위치순서대로 기술하는 것으로 한다.
'   << 공통변수 정의 부분 >>
'   공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은
'               통일하도록 한다.
'   1. 공통컨트롤을 Call하는 변수
'          ADF (ADS, ADC, ADF는 그대로 사용)
'          - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
'   2. 공통컨트롤에서 Return된 값을 받는 변수
'           strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'   설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery()

    Dim IntRetCD

    FncQuery = False                                                        '⊙: Processing is NG

    Err.Clear                                                               '☜: Protect system from crashing

    Dim firstCheck
    Dim secondCheck
    firstCheck = "N"
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        firstCheck = "Y"
    End If

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        secondCheck ="Y"
    End If

    If firstCheck = "Y" Or secondCheck = "Y" Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")   '⊙: Display Message(There is no changed data.)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If


    If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdToDt) = False Then Exit Function

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")                                  '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      '⊙: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then                             '⊙: This function check indispensable field
       Exit Function
    End If

    frm1.hPlantCd.value     = frm1.txtPlantCd.value
    frm1.hItemCd.value      = frm1.txtItemCd.value
    frm1.hProdOrderNo.value = frm1.txtProdOrderNo.value
    frm1.hProdFromDt.value  = frm1.txtProdFromDt.Text
    frm1.hProdToDt.value    = frm1.txtProdToDt.Text
    frm1.hOrderType.value   = frm1.cboOrderType.value
    frm1.hTrackingNo.value  = frm1.txtTrackingNo.value

    If frm1.rdoStatus2.checked = True Then
        frm1.hOrderStatus.value = "OP"
    Else
        frm1.hOrderStatus.value = "RL"
    End If

	lgFlgQueryCnt = 0

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function                                                           '☜: Query db data

    FncQuery = True                                                         '⊙: Processing is OK

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

    FncSave = False                                             '⊙: Processing is NG

    Err.Clear                                                   '☜: Protect system from crashing


    Dim firstCheck
    Dim secondCheck
    firstCheck = "N"
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        firstCheck = "Y"

        If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
           Exit Function
        End If

        Call DbSave()
    End If

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        secondCheck = "Y"

        If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
           Exit Function
        End If

        For LngRows = 1 To frm1.vspdData2.MaxRows
            frm1.vspdData2.Row = LngRows
            frm1.vspdData2.Col = C_ReqQty2
            If frm1.vspdData2.Value <= 0 Then
                Call DisplayMsgBox("189506", "x", "x", "x")
                Call SheetFocus2(LngRows, C_ReqQty2)
                Exit Function
            End If    
        Next    

        Call DbSave2()
    End If

    If firstCheck = "N" AND secondCheck = "N" Then
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")   '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
'    If DbSave = False Then Exit Function                        '☜: Save db data

    FncSave = True                                              '⊙: Processing is OK
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy()

    Dim LngRow

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    frm1.vspdData.focus
    Set gActiveElement = document.activeElement
    frm1.vspdData.EditMode = True

    frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData

    ggoSpread.CopyRow
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_ProdtOrderNo
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_PlanOrderNo
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_TrackingNo
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ItemGroupCd
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ItemGroupNm
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_MRPRunNo
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ParentOrderNo
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ParentOprNo
    frm1.vspdData.Text = ""

    frm1.vspdData.ReDraw = True

    LngRow = frm1.vspdData.ActiveRow
    SetSpreadColor LngRow, LngRow
    frm1.vspdData.Row = LngRow
    frm1.vspdData.Col = C_OrderType
    frm1.vspdData.Text = "2"
    frm1.vspdData.Col = C_OrderTypeDesc
    frm1.vspdData.Text = "2"
    frm1.vspdData.Col = C_ReWorkFlag
    frm1.vspdData.Text = "N"

    frm1.vspdData.Col = C_ItemCode
    frm1.vspdData.Row = LngRow
    Call LookUpItemByPlant(frm1.vspdData.Text, LngRow)

    Call InitData(LngRow)

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
    If whereVspdData = "TOP" Or whereVspdData = "" Then
        If frm1.vspdData.MaxRows < 1 Then Exit Function

        ggoSpread.Source = frm1.vspdData
        ggoSpread.EditUndo                                                  '☜: Protect system from crashing

        Call initData(frm1.vspdData.ActiveRow)

    End If

    If whereVspdData = "DOWN" Then
        If frm1.vspdData2.MaxRows < 1 Then Exit Function

        ggoSpread.Source = frm1.vspdData2
        ggoSpread.EditUndo                                                  '☜: Protect system from crashing

         frm1.vspdData2.Redraw = True
    End If
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim pvRow

    Dim slCd, slNm

    If whereVspdData = "TOP" Or whereVspdData = "" then
        On Error Resume Next

        FncInsertRow = False

        If IsNumeric(Trim(pvRowCnt)) Then
            imRow = Cint(pvRowCnt)
        Else
            imRow = AskSpdSheetAddRowCount()

            If imRow = "" Then
                Exit Function
            End If
        End If

        With frm1
            frm1.vspdData.focus

            Set gActiveElement = document.activeElement

            ggoSpread.Source = frm1.vspdData

            frm1.vspdData.ReDraw = False

            ggoSpread.InsertRow frm1.vspdData.ActiveRow, imRow

            SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow + imRow -1

            frm1.vspdData.Col = C_ReWorkFlag

            For pvRow = frm1.vspdData.ActiveRow To frm1.vspdData.ActiveRow + imRow -1
                frm1.vspdData.Row = pvRow
                frm1.vspdData.Text = "N"
            Next

            frm1.vspdData.ReDraw = True
        End With

        Set gActiveElement = document.ActiveElement

        If Err.number = 0 Then FncInsertRow = True
    End If

    If whereVspdData = "DOWN" Then
        On Error Resume Next

        ggoSpread.Source = frm1.vspdData

        frm1.vspdData.Row = frm1.vspdData.ActiveRow
        frm1.vspdData.Col = C_WcCd
        slCd = frm1.vspdData.Text
        frm1.vspdData.Col = C_WcNm
        slNm = frm1.vspdData.Text

        ggoSpread.Source = frm1.vspdData2

        FncInsertRow = False

        If IsNumeric(Trim(pvRowCnt)) Then
            imRow = Cint(pvRowCnt)
        Else
            imRow = AskSpdSheetAddRowCount()

            If imRow = "" Then
                Exit Function
            End If
        End If

        With frm1

            frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
            frm1.vspdData2.Col = C_OrderStatus2

            If frm1.vspdData2.Text = "ST" Then
                Call DisplayMsgBox("189520", "x", "x", "x")
                Exit Function
            End IF

            If frm1.vspdData2.Text = "CL" Then
                Call DisplayMsgBox("189523", "x", "x", "x")
                Exit Function
            End IF

            frm1.vspdData2.focus

            Set gActiveElement = document.activeElement

            frm1.vspdData2.ReDraw = False

            ggoSpread.InsertRow frm1.vspdData2.ActiveRow, imRow

            frm1.vspdData2.Row = frm1.vspdData2.ActiveRow

            Dim C_OprNo2Insert  
            Dim C_JobCd2Insert  
            Dim C_JobDesc2Insert
            Dim C_WcCd2Insert   
            Dim C_WcNm2Insert
            Dim C_ProdtOrderNo2Insert

            pvRow = frm1.vspdData2.ActiveRow-1

            frm1.vspdData2.Row = pvRow

            frm1.vspdData2.Col = C_OprNo2
            C_OprNo2Insert = frm1.vspdData2.text
            frm1.vspdData2.Col = C_JobCd2
            C_JobCd2Insert = frm1.vspdData2.text
            frm1.vspdData2.Col = C_JobDesc2
            C_JobDesc2Insert = frm1.vspdData2.text
            frm1.vspdData2.Col = C_WcCd2
            C_WcCd2Insert = frm1.vspdData2.text
            frm1.vspdData2.Col = C_WcNm2
            C_WcNm2Insert = frm1.vspdData2.text
            frm1.vspdData2.Col = C_ProdtOrderNo2
            C_ProdtOrderNo2Insert = frm1.vspdData2.text

            pvRow = frm1.vspdData2.ActiveRow

            frm1.vspdData2.Row = pvRow


            frm1.vspdData2.Col = C_OprNo2
            frm1.vspdData2.value = C_OprNo2Insert
            frm1.vspdData2.Col = C_JobCd2
            For pvRow = frm1.vspdData2.ActiveRow To frm1.vspdData2.ActiveRow + imRow -1
                frm1.vspdData2.Row = pvRow
                frm1.vspdData2.Text = C_JobCd2Insert
            Next
            frm1.vspdData2.Col = C_JobDesc2
            For pvRow = frm1.vspdData2.ActiveRow To frm1.vspdData2.ActiveRow + imRow -1
                frm1.vspdData2.Row = pvRow
                frm1.vspdData2.Text = C_JobDesc2Insert
            Next
            frm1.vspdData2.Col = C_WcCd2
            frm1.vspdData2.value = C_WcCd2Insert
            frm1.vspdData2.Col = C_WcNm2
            frm1.vspdData2.value = C_WcNm2Insert

            frm1.vspdData2.Col = C_ProdtOrderNo2
            frm1.vspdData2.value = C_ProdtOrderNo2Insert

            frm1.vspdData2.Col = C_SlCd2
            frm1.vspdData2.value = slCd
            frm1.vspdData2.Col = C_SlNm2
            frm1.vspdData2.value = slNm

            SetSpreadColor2 frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow + imRow -1

            frm1.vspdData2.ReDraw = True

            Set gActiveElement = document.ActiveElement

            If Err.number = 0 Then FncInsertRow = True

        End With
    End If

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
    Dim lDelRows
    Dim iDelRowCnt, i

    If whereVspdData = "TOP" Or whereVspdData = "" Then
        If frm1.vspdData.MaxRows < 1 Then Exit Function

        ggoSpread.Source = frm1.vspdData

        lDelRows = ggoSpread.DeleteRow
        lgLngCurRows = lDelRows + lgLngCurRows
    End If
    
    If whereVspdData = "DOWN" Then
        With frm1
            ggoSpread.Source = frm1.vspdData2

            frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
            frm1.vspdData2.Col = C_OrderStatus2

            If frm1.vspdData2.Text = "ST" Then
                Call DisplayMsgBox("189520", "x", "x", "x")
                Exit Function
            End IF    
       
            If frm1.vspdData2.Text = "CL" Then
                Call DisplayMsgBox("189523", "x", "x", "x")
                Exit Function
            End IF    
       
            If frm1.vspdData2.MaxRows < 1 Then Exit Function

        End With

        ggoSpread.Source = frm1.vspdData2
        lDelRows = ggoSpread.DeleteRow
        lgLngCurRows = lDelRows + lgLngCurRows
    End If
    
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
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext()
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)
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

    Dim firstCheck
    Dim secondCheck
    firstCheck = "N"
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        firstCheck = "Y"
    End If

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        secondCheck ="Y"
    End If

    If firstCheck = "Y" Or secondCheck = "Y" Then
        IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")   '⊙: Display Message(There is no changed data.)

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True

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


Function btnRelease_onClick()

	If lgButtonSelection = "Release" Then
        Call ReleaseOrder
	Else ' Cancel
        Call CancelOrder
	End If

End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'   설명 :
'******************************************************************************************************%>
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False

    lgFlgQueryCnt = lgFlgQueryCnt + 1

    If LayerShowHide(1) = False Then Exit Function

    Dim strVal

    If lgIntFlgMode = parent.OPMD_UMODE Then
        strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
        strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
        strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)
        strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
        strVal = strVal & "&txtStartDt=" & Trim(frm1.hProdFromDt.value)
        strVal = strVal & "&txtEndDt=" & Trim(frm1.hProdToDt.value)
        strVal = strVal & "&cboOrderType=" & Trim(frm1.hOrderType.value)
        strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
        strVal = strVal & "&txtMRPRunNo=" & Trim(frm1.hMRPRunNo.value)
        strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
        strVal = strVal & "&rdoOrderStatus=" & Trim(frm1.hOrderStatus.value)
	    strVal = strVal & "&txtBsitemcd=" & Trim(frm1.txtBsitemcd.value)
        strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    Else
        strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
        strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
        strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)
        strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
        strVal = strVal & "&txtStartDt=" & Trim(frm1.txtProdFromDt.text)
        strVal = strVal & "&txtEndDt=" & Trim(frm1.txtProdToDt.text)
        strVal = strVal & "&cboOrderType=" & Trim(frm1.cboOrderType.value)
        strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
        strVal = strVal & "&txtMRPRunNo=" & Trim(frm1.txtMRPRunNo.value)
        strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
        strVal = strVal & "&rdoOrderStatus=" & Trim(frm1.hOrderStatus.value)
	    strVal = strVal & "&txtBsitemcd=" & Trim(frm1.txtBsitemcd.value)
        strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
    End If

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)             '☆: 조회 성공후 실행로직

    Dim lRow
    Dim LngRow


    Call ggoOper.LockField(Document, "Q")                                       '⊙: This function lock the suitable field
    Call SetToolBar("11001111001111")                                           '⊙: 버튼 툴바 제어

    frm1.vspdData.ReDraw = False

    ggoSpread.SSSetProtected C_ItemCode, LngMaxRow, frm1.vspdData.MaxRows
    ggoSpread.SSSetProtected C_ItemPopup, LngMaxRow, frm1.vspdData.MaxRows
    ggoSpread.SSSetProtected C_ProdtOrderNo, LngMaxRow, frm1.vspdData.MaxRows
    ggoSpread.SSSetProtected C_ReWorkFlag, LngMaxRow, frm1.vspdData.MaxRows

    With frm1.vspdData

        If .MaxRows < 1 Then Exit Function

        Call InitData(LngMaxRow)

        Call InitTrackingNCost("A")

        If lgIntFlgMode = parent.OPMD_CMODE Then

            .Row = 1
            .Col = C_OrderUnitMFG
            frm1.txtOrderUnitMFG.value = .Text
            .Col = C_MinMRPQty
            frm1.txtMinMRPQty.value = .Text
            .Col = C_FixedMRPQty
            frm1.txtFixedMRPQty.value = .Text
            .Col = C_MaxMRPQty
            frm1.txtMaxMRPQty.value = .Text
            .Col = C_RoundQty
            frm1.txtRoundQty.value = .Text
            .Col = C_ValidFromDT
            frm1.txtValidFromDT.Text = .Text
            .Col = C_ValidToDT
            frm1.txtValidToDT.Text = .Text
            .Col = C_OrderLtMFG
            frm1.txtOrderLtMFG.value = .Text
            .Col = C_ScrapRateMFG
            frm1.txtScrapRateMFG.value = .Text
            .Col = C_MPSMgr
            frm1.txtMPSMgr.value = .Text
            .Col = C_MRPMgr
            frm1.txtMRPMgr.value = .Text
            .Col = C_ProdMgr
            frm1.txtProdMgr.value = .Text

            Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
            Set gActiveElement = document.activeElement

        End If

    End With

    frm1.vspdData.ReDraw = True
    '-----------------------
    'Reset variables area
    '-----------------------

    If frm1.rdoStatus1.checked = True Then
        ggoSpread.SpreadLock  C_OrderQty, -1, C_OrderQty
        ggoSpread.SpreadLock  C_BATCHCNT, -1, C_BATCHCNT        '20080307
        ggoSpread.SpreadLock  C_OrderUnit, -1, C_OrderUnit
        ggoSpread.SpreadLock  C_SLCd, -1, C_SLCd
        ggoSpread.SpreadLock  C_RoutingNo, -1, C_RoutingNo
        ggoSpread.SpreadLock  C_PlanStartDt, -1, C_PlanStartDt
        ggoSpread.SpreadLock  C_PlanEndDt, -1, C_PlanEndDt
        ggoSpread.SpreadLock  C_WcCd, -1, C_WcCd
        ggoSpread.SpreadLock  C_Remark, -1, C_Remark

'        frm1.btnRelease.disabled = True
    	frm1.btnRelease.value = "제조오더확정취소"
        lgButtonSelection = "Cancel"
    Else
        ggoSpread.SpreadUnLock  C_OrderQty, -1, C_OrderQty
        ggoSpread.SpreadUnLock  C_BATCHCNT, -1, C_BATCHCNT        '20080307
        ggoSpread.SpreadUnLock  C_OrderUnit, -1, C_OrderUnit
        ggoSpread.SpreadUnLock  C_SLCd, -1, C_SLCd
        ggoSpread.SpreadUnLock  C_RoutingNo, -1, C_RoutingNo
        ggoSpread.SpreadUnLock  C_PlanStartDt, -1, C_PlanStartDt
        ggoSpread.SpreadUnLock  C_PlanEndDt, -1, C_PlanEndDt
        ggoSpread.SpreadUnLock  C_WcCd, -1, C_WcCd
        ggoSpread.SpreadUnLock  C_Remark, -1, C_Remark

        ggoSpread.SSSetRequired  C_OrderQty, -1
        ggoSpread.SSSetRequired  C_BATCHCNT, -1     '20080307

        ggoSpread.SSSetRequired  C_OrderUnit, -1
        ggoSpread.SSSetRequired  C_SLCd, -1
        ggoSpread.SSSetRequired  C_RoutingNo, -1
        ggoSpread.SSSetRequired  C_PlanStartDt, -1
        ggoSpread.SSSetRequired  C_PlanEndDt, -1
'20080227::hanc        ggoSpread.SSSetRequired  C_WcCd, -1

'        frm1.btnRelease.disabled = False
    	frm1.btnRelease.value = "제조오더확정"
        lgButtonSelection = "Release"
    End If

    frm1.btnRelease.disabled = False

	frm1.vspdData.Col = 1
	frm1.vspdData.Row = 1

	lgOldRow = 1

	If lgFlgQueryCnt = 1 Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
			Set gActiveElement = document.activeElement
			Call DbDtlQuery
		End If
	End If

	lgIntFlgMode = parent.OPMD_UMODE	                                                 '⊙: Indicates that current mode is Update mode

End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery가 성공적이지 아닐경우
'========================================================================================
Function DbQueryNotOk()

    Call SetToolBar("11001101001111")                                                       '⊙: 버튼 툴바 제어

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_CMODE                                                    '⊙: Indicates that current mode is Update mode

End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave()

    Dim IntRows
    Dim strVal, strDel
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
    Dim iProdtOrderQty, iBatchCnt

    Dim lGrpcnt

    ggoSpread.Source = frm1.vspdData                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then
        lGrpCnt = 1

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

        DbSave = False                                                          '⊙: Processing is NG

        If LayerShowHide(1) = False Then Exit Function


        frm1.txtMode.value = parent.UID_M0002                                       '☜: 저장 상태
        frm1.txtFlgMode.value = lgIntFlgMode                                    '☜: 신규입력/수정 상태

        '-----------------------
        'Data manipulate area
        '-----------------------
        lGrpCnt = 1

        With frm1.vspdData

            For IntRows = 1 To .MaxRows

                .Row = IntRows
                .Col = 0

                Select Case .Text

                    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

                        strVal = ""

                        If .Text = ggoSpread.InsertFlag Then
                            strVal = strVal & "CREATE" & iColSep                '⊙: C=Create, Sheet가 2개 이므로 구별
                        Else
                            strVal = strVal & "UPDATE" & iColSep                '⊙: U=Update
                        End If

                        ' 2. Plant Code
                        strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
                        ' 3. Production Order No.
                        .Col = C_ProdtOrderNo
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 4. Item Code
                        .Col = C_ItemCode
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 5. Re-Work Flag
                        .Col = C_ReWorkFlag
                        strVal = strVal & Trim(.Text) & iColSep

                        .Col = 0
                        If .Text = ggoSpread.InsertFlag Then            'insert
                            ' 6. Order Quantity
                            .Col = C_OrderQty
                            strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
                        Else                                            'update
                            .Col = C_OrderQty
                            iProdtOrderQty  =   CDbl(.Text)

                            .Col = C_BATCHCNT
                            iBatchCnt = iProdtOrderQty * CDbl(.Text)

                            strVal = strVal & iBatchCnt & iColSep          '22
                            
                        End If

                        ' 7. Order Unit
                        .Col = C_OrderUnit
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 8. Order Quantity In Base Unit
                        strVal = strVal & UNIConvNum("0",0) & iColSep
                        ' 9. Base Unit
                        .Col = C_BaseUnit
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 10. S/L Code
                        .Col = C_SLCD
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 11. Routing No.
                        .Col = C_RoutingNo
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 12. Planned Start Date
                        .Col = C_PlanStartDt
                        strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
                        ' 13. Planned End Date
                        .Col = C_PlanEndDt
                        strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
                        ' 14. BOM Type
                        .Col = C_BOMNo
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 15. Tracking No.
                        .Col = C_TrackingNo
                        If Trim(.Text) = "" Then
                            strVal = strVal & "*" & iColSep                             '☆: Tracking No.
                        Else
                            strVal = strVal & UCase(Trim(.Text)) & iColSep      '☆: Tracking No.
                        End If
                        ' 16. User ID
                        .Col = C_Remark
                        ' 17. remark
                        strVal = strVal & Trim(.Text) & iColSep
                        ' 18. Parent Order No
                        .Col = C_ParentOrderNo
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 19. Parent Opr No
                        .Col = C_ParentOprNo
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 20. CostCd
                        .Col = C_CostCd
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 21. Row Count
                        strVal = strVal & IntRows & iColSep
                        ' 22. opr no
                        .Col = C_OprNo
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' 23. wc cd
                        .Col = C_WcCd
                        strVal = strVal & UCase(Trim(.Text)) & iColSep

                        ' 20080310::hanc 24. C_BATCHCNT
                        .Col = C_BATCHCNT
                        strVal = strVal & UCase(Trim(.Text)) & iColSep          '22

                        ' 6. Order Quantity
                        .Col = C_OrderQty
                        strVal = strVal & UNIConvNum(Trim(.Text),0) & iRowSep       '23 20080310::hanc:: 이걸 하는 이유  : 오더수량 * 배치cnt의 값을 p_reservation에 적용하기 위해서 ui단에서 값을 던져주고
                                                                                                                           '그리고 P_PRODUCTION_ORDER_HEADER에는 23번째 변수를 적용한다.
                    

                        lGrpCnt = lGrpCnt + 1

                    Case ggoSpread.DeleteFlag

                        strDel = ""

                        strDel = strDel & "DELETE" & iColSep                '⊙: D=Delete
                        ' 2. Plant Code
                        strDel = strDel & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
                        ' 3. Production Order No.
                        .Col = C_ProdtOrderNo
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' 4. Item Code
                        .Col = C_ItemCode
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' 5. Re-Work Flag
                        .Col = C_ReWorkFlag
                        strDel = strDel & Trim(.Text) & iColSep
                        ' 6. Order Quantity
                        .Col = C_OrderQty
                        strDel = strDel & UNIConvNum(Trim(.Text),0) & iColSep
                        ' 7. Order Unit
                        .Col = C_OrderUnit
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' 8. Order Quantity In Base Unit
                        .Col = C_OrderQtyInBaseUnit
                        strDel = strDel & UNIConvNum(Trim(.Text),0) & iColSep
                        ' 9. Base Unit
                        .Col = C_BaseUnit
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' 10. S/L Code
                        .Col = C_SLCD
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' 11. Routing No.
                        .Col = C_RoutingNo
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' 12. Planned Start Date
                        .Col = C_PlanStartDt
                        strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep
                        ' 13. Planned End Date
                        .Col = C_PlanEndDt
                        strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep
                        ' 14. BOM Type
                        .Col = C_BOMNo
                        strDel = strDel & Trim(.Text) & iColSep
                        ' 15. Tracking No.
                        .Col = C_TrackingNo
                        If Trim(.Text) = "" Then
                            strDel = strDel & "*" & iColSep                             '☆: Tracking No.
                        Else
                            strDel = strDel & UCase(Trim(.Text)) & iColSep      '☆: Tracking No.
                        End If
                        ' 16. User ID
                        ' 17. remark
                        strDel = strDel & Trim(.Text) & iColSep
                        ' 18. Parent Order No
                        strDel = strDel & "" & iColSep
                        ' 19. Parent Opr No
                        strDel = strDel & "" & iColSep
                        ' 20. Cost Cd
                        strDel = strDel & "" & iColSep
                        ' 21. Row Count
                        strDel = strDel & IntRows & iColSep
                        ' 22. opr no
                        .Col = C_OprNo
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' 23. wc cd
                        .Col = C_WcCd
                        strDel = strDel & UCase(Trim(.Text)) & iRowSep

                        lGrpCnt = lGrpCnt + 1

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

        frm1.txtMaxRows.value = lGrpCnt-1
        Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)                                '☜: 비지니스 ASP 를 가동

        DbSave = True                                                           ' ⊙: Processing is OK

    End If

End Function

Function DbSave2()
    Dim IntRows
    Dim strVal, strDel
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

    Dim lGrpcnt

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged

        DbSave2 = False                                                          '⊙: Processing is NG
        
        Call LayerShowHide(1)

        With frm1
            .txtMode.value = parent.UID_M0002											'☜: 저장 상태 
            .txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
            .txtUpdtUserId.value = parent.gUsrID
            .txtInsrtUserId.value  = parent.gUsrID
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

        With frm1.vspdData2

            For IntRows = 1 To .MaxRows
        
                .Row = IntRows
                .Col = 0
                
                Select Case .Text
                
                    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                        
                        strVal = ""
                        
                        If .Text = ggoSpread.InsertFlag Then
                            strVal = strVal & "CREATE" & iColSep				'⊙: C=Create, Sheet가 2개 이므로 구별 
                        Else
                            strVal = strVal & "UPDATE" & iColSep				'⊙: U=Update
                        End If
                                        
                        strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
                        .Col = C_ProdtOrderNo2	' Production Order No.
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' Plan Order No.
                        strVal = strVal & iColSep
                        .Col = C_OprNo2			' Opr No.
                        strVal = strVal & Trim(.Text) & iColSep
                        .Col = C_Seq2		' Sequence
                        strVal = strVal & Trim(.Text) & iColSep
                        ' Resvrd Status
                        strVal = strVal & iColSep
                        .Col = C_ReqDt2		' Required Date
                        strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
                        .Col = C_ReqQty2		' Required Quantity
                        strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep
                        .Col = C_ReqNo2			'  Required No.
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        ' Resvrd Type
                        strVal = strVal & iColSep
                        .Col = C_TrackingNo2	' Tracking No.
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        .Col = C_BaseUnit2			' Base Unit
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        .Col = C_IssueMthd2		' Issue Method
                        strVal = strVal & Trim(.Text) & iColSep
                        .Col = C_ItemCd2		' Child Item Cd
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        .Col = C_SlCd2		'  Storage Location
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        .Col = C_WcCd2			'  Work Center
                        strVal = strVal & UCase(Trim(.Text)) & iColSep
                        'Row Count
                        strVal = strVal & IntRows & parent.gRowSep

                    Case ggoSpread.DeleteFlag
                        
                        strDel = ""
                        strDel = strDel & "DELETE" & iColSep				'⊙: D=Delete
                        strDel = strDel & Trim(frm1.txtPlantCd.value) & iColSep
                        .Col = C_ProdtOrderNo2  ' Production Order No.
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' Plan Order No.
                        strDel = strDel & iColSep
                        .Col = C_OprNo2         ' Opr No.
                        strDel = strDel & Trim(.Text) & iColSep
                        .Col = C_Seq2 		' Sequence
                        strDel = strDel & Trim(.Text) & iColSep
                        ' Resvrd Status
                        strDel = strDel & iColSep
                        .Col = C_ReqDt2         ' Required Date
                        strDel = strDel & UNIConvDate(Trim(.Text)) & iColSep
                        .Col = C_ReqQty2        ' Required Quantity
                        strDel = strDel & UNIConvNum(Trim(.Text),0) & iColSep
                        .Col = C_ReqNo2         '  Required No.
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        ' Resvrd Type
                        strDel = strDel & iColSep
                        .Col = C_TrackingNo2     'Tracking No.
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        .Col = C_BaseUnit2      ' Base Unit
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        .Col = C_IssueMthd2     ' Issue Method
                        strDel = strDel & Trim(.Text) & iColSep
                        .Col = C_ItemCd2        ' Child Item Cd
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        .Col = C_SlCd2          '  Storage Location
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        .Col = C_WcCd2          '  Work Center
                        strDel = strDel & UCase(Trim(.Text)) & iColSep
                        'Row Count
                        strDel = strDel & IntRows & parent.gRowSep

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
           objTEXTAREA.name   = "txtCUSpread2"
           objTEXTAREA.value = Join(iTmpCUBuffer,"")
           
           divTextArea.appendChild(objTEXTAREA)
        End If   

        If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
           Set objTEXTAREA = document.createElement("TEXTAREA")
           objTEXTAREA.name = "txtDSpread2"
           objTEXTAREA.value = Join(iTmpDBuffer,"")
           divTextArea.appendChild(objTEXTAREA)     
        End If   

        Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID2)								'☜: 저장 비지니스 ASP 를 가동 

        DbSave2 = True                                                           ' ⊙: Processing is OK

    End If
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김
'========================================================================================

Function DbSaveOk()                                                 '☆: 저장 성공후 실행 로직
    Dim secondCheck
    secondCheck ="N"

    ggoSpread.Source = frm1.vspdData2                        '⊙: Preset spreadsheet pointer
    If ggoSpread.SSCheckChange = True Then                  '⊙: Check If data is chaged
        Exit Function
    End If

    Call InitVariables

    ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    
    ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0

    Call RemovedivTextArea
    Call MainQuery()
End Function

Function DbSaveOk2()                                                 '☆: 저장 성공후 실행 로직
    Call InitVariables

    ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    
    ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0

    Call RemovedivTextArea
    Call MainQuery()

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

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌
'==============================================================================
Function SheetFocus(lRow, lCol)
    ggoSpread.source = frm1.vspdData
    frm1.vspdData.focus
'    frm1.vspdData.Row = lRow
    frm1.vspdData.Col = lCol
    frm1.vspdData.Action = 0
    frm1.vspdData.SelStart = 0
    frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function

Function SheetFocus2(lRow, lCol)
    ggoSpread.source = frm1.vspdData2
    frm1.vspdData2.focus
    frm1.vspdData2.Row = lRow
    frm1.vspdData2.Col = lCol
    frm1.vspdData2.Action = 0
    frm1.vspdData2.SelStart = 0
    frm1.vspdData2.SelLength = len(frm1.vspdData2.Text)
End Function
'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery()
    If frm1.rdoStatus1.checked = True Then
        
        If IntRetCD = vbNo Then
            Exit Function
        End If

        Dim strVal
        Dim boolExist
        Dim lngRows
        Dim strOprCd

'        Call SetToolBar("11000000000111")

        boolExist = False
        With frm1
            .vspdData.Row = .vspdData.ActiveRow
            .vspdData.Col = C_ProdtOrderNo
            strOprCd = .vspdData.Text

            DbDtlQuery = False

            .vspdData.Row = .vspdData.ActiveRow

            Call LayerShowHide(1)

            If lgIntFlgMode = parent.OPMD_UMODE Then
                strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜:
                strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타
                strVal = strVal & "&txtProdtOrderNo=" & Trim(strOprCd)
                .vspdData.Col = C_ProdtOrderNo
                strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
            Else
                strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜:
                strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타
                strVal = strVal & "&txtProdtOrderNo=" & Trim(strOprCd)
                .vspdData.Col = C_ProdtOrderNo
                strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows

            End If


            Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동

        End With

        DbDtlQuery = True
    End If
End Function

Function DbDtlQueryOk(ByVal LngMaxRow)												'☆: 조회 성공후 실행로직
	Dim LngRow

	frm1.vspdData2.Col = C_InsideFlag2

    If frm1.vspdData2.Text = "N" Then
		Call SetToolBar("11001001000111")										'⊙: 버튼 툴바 제어
	Else
		Call SetToolBar("11001111000111")										'⊙: 버튼 툴바 제어
	End IF

    '-----------------------
    'Reset variables area
    '-----------------------
	frm1.vspdData2.ReDraw = False

	ggoSpread.Source = frm1.vspdData2

	ggoSpread.SSSetProtected C_ItemCd2,		1, frm1.vspdData2.MaxRows
	ggoSpread.SSSetProtected C_ItemCdPopup2,	1, frm1.vspdData2.MaxRows

	With frm1.vspdData2

		For LngRow = 1 To .MaxRows

			.Row = LngRow
			.Col = C_OrderStatus2

			If .Text = "CL" Then
				ggoSpread.SSSetProtected C_ReqQty2,			LngRow, LngRow
				ggoSpread.SSSetProtected C_ReqDt2,			LngRow, LngRow
    			ggoSpread.SSSetProtected C_SlCd2,		LngRow, LngRow
				ggoSpread.SpreadLock C_SlCdPopup2,		LngRow, C_SlCdPopup2, LngRow
			Else
				ggoSpread.SSSetRequired C_ReqQty2,			LngRow, LngRow
				ggoSpread.SSSetRequired C_ReqDt2,			LngRow, LngRow
				ggoSpread.SSSetRequired C_SlCd2,		LngRow, LngRow
				ggoSpread.SpreadUnLock C_SlCdPopup2,	LngRow, C_SlCdPopup2, LngRow
			End If

		Next

	End With

	lgAfterQryFlg = True

	frm1.vspdData2.ReDraw = True

End Function



'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
	
	Dim strItemCd
	Dim strHndItemCd, strHndOprNo
	Dim i
	Dim strReqDt, strEndDt
	Dim	DblRqrdQty, DblIssuedQty
	Dim lNewRow, lOldRow

'	lOldRow = frm1.vspdData1.ActiveRow
					
	With frm1.vspdData2

		Select Case Col

		    Case C_ItemCd2

				.Row = Row
				.Col = C_ItemCd2
				strItemCd = .Text
				
				If strItemCd = "" Then Exit Sub
				
				For i = 1 To .MaxRows
					If i <> Row Then
						.Row = i
						.Col = C_ItemCd2
						If UCase(Trim(.Text)) = UCase(Trim(strItemCd)) Then
							Call DisplayMsgBox("189504", "x", "x", "x")
							.Row = Row
							.Text = ""
							Exit Sub
						End If
					End If						
				Next
				
				Call LookUpItemByPlant2(strItemCd, Row)
 
		    Case C_ReqDt2
				
				' 필요일이 공정의 완료예정일 보다 미래일 수 없다.
				.Row = Row
				.Col = C_ReqDt2
				strReqDt = .Text
				.Col = C_OprNo2

                
                frm1.vspdData2.Col = C_ProdtOrderNo2
                frm1.vspdData2.Row = Row-1
                
                strSelect = " Plan_Start_Dt, Plan_Compt_Dt "

                If  CommonQueryRs2by2(strSelect, " p_production_order_header ", " prodt_order_no = "& FilterVar(Frm1.vspdData2.Text, "''", "S"), lgF0) = False Then
                    Call DisplayMsgBox("189505", "x", "x", "x") '필요일이 투입공정의 완료예정일보다 미래일 수 없습니다.
                End If

	            lgF0 = Split(lgF0, Chr(11))

                strEndDt = lgF0(2)

				If UniConvDateAToB(strReqDt, parent.gDateFormat, parent.gServerDateFormat) > UniConvDateAToB(strEndDt, parent.gDateFormat, parent.gServerDateFormat) Then  
					Call DisplayMsgBox("189505", "x", "x", "x") '필요일이 투입공정의 완료예정일보다 미래일 수 없습니다.
				End If

				.Col = C_ItemCd2
			
				If .Text <> "" Then
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.UpdateRow Row
				End If
				
		    Case C_ReqQty2

				.Row = Row
				.Col = C_OprNo2
				strHndOprNo = .Text
				.Col = C_ReqQty2
				DblRqrdQty = .Text
				.Col = C_IssuedQty2

				If UNICDbl(DblRqrdQty) < UNICDbl(.Text) Then  
					Call DisplayMsgBox("189521", "x", "x", "x")  '부품 필요량을 출고량보다 적게 변경할 수 없습니다.
				End If

				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row

		    Case C_SlCd2

                frm1.vspdData2.Col = C_SlCd2
                Call TabSLCd(.Value, "DOWN", Row)
				
		End Select

	End With

End Sub




Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

Dim strCode
Dim strName

    With frm1.vspdData2
    
		ggoSpread.Source = frm1.vspdData2
		If Row < 1 Then Exit Sub

		Select Case Col

		    Case C_ItemCdPopup2
				.Col = C_ItemCd2
				.Row = Row
				strCode = .Text
				.Col = C_ItemNm2
				.Row = Row
				strName = .Text
				Call OpenItemInfo3(strCode, strName, Row)
				Call SetActiveCell(frm1.vspdData2, C_ItemCd2, Row, "M","X","X")
				Set gActiveElement = document.activeElement
	    
		    Case C_SlCdPopup2
				.Col = C_SlCd2
				.Row = Row
				strCode = .Text
				.Col = C_SlNm2
				.Row = Row
				strName = .Text
				Call OpenSLCD2(strCode, strName, Row)
				Call SetActiveCell(frm1.vspdData2, C_MajorSLCd, Row, "M","X","X")
				
		End Select

	End With

End Sub

Function OpenItemInfo3(Byval strCode, Byval strName, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 'ITEM_CD					' Field명(0)
	arrField(1) = 2 'ITEM_NM					' Field명(1)
	arrField(2) = 3	'SPECIFICATION
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo3(arrRet, Row)
	End If	
	
	Call SetFocusToDocument("M")

End Function


'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo3(Byval arrRet, Byval Row)

	Dim i

    With frm1.vspddata2

		For i = 1 to .MaxRows
			.Row = i
			.Col = C_ItemCd2
			If .Text = arrRet(0) Then
				Call DisplayMsgBox("189504", "x", "x", "x")
				Exit Function
			End If
		Next
		
		.Row = Row
		.Col = C_ItemCd2		
		.Text = arrRet(0)
		.Col = C_ItemNm2
		.Text = arrRet(1)
		.Col = C_Spec2
		.Text = arrRet(2)

		Call vspdData2_Change(C_ItemCd2,  Row)

    End With

End Function



'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'	Name : LookUpItem ByPlant()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 

Function LookUpItemByPlant2(Byval StrItemCd, Byval Row)
    
	Dim strVal
	Dim strSelect, strWhere
	Dim gComNum1000, gComNumDec, gAPNum1000, gAPNumDec
	
	gComNum1000 = parent.gComNum1000
	gComNumDec = parent.gComNumDec
	gAPNum1000 = parent.gAPNum1000
	gAPNumDec = parent.gAPNumDec

	If strItemCd = "" Then Exit Function
	frm1.vspdData2.Col = C_ProdtOrderNo2
	frm1.vspdData2.Row = Row-1

	strSelect = " A.ITEM_CD, A.BASIC_UNIT, A.ITEM_NM, A.SPEC, A.PHANTOM_FLG, B.VALID_FLG ITEM_VALID_FLG, B.PROCUR_TYPE,  "
	strSelect = strSelect & " B.VALID_FLG PLANT_VALID_FLG,   B.TRACKING_FLG, B.ORDER_UNIT_MFG, B.ORDER_LT_MFG,B.ISSUED_SL_CD, C.SL_NM, "
	strSelect = strSelect & " B.ISSUE_MTHD,   DBO.UFN_GETCODENAME( " & FilterVar("P1016", "''", "S") & " , B.ISSUE_MTHD ) AS  ISSUE_DESC  "
	strSelect = strSelect & ", (select tracking_no from p_production_order_header where prodt_order_no = " & FilterVar(Frm1.vspdData2.Text, "''", "S") & ") as tracking_no_q"

    frm1.vspdData2.Col = C_ItemCd2
	frm1.vspdData2.Row = Row

	strWhere = " A.ITEM_CD = B.ITEM_CD       AND B.ISSUED_SL_CD = C.SL_CD       AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	strWhere = strWhere & " AND B.ITEM_CD = " & FilterVar(Frm1.vspdData2.Text, "''", "S")
	
	If 	CommonQueryRs2by2(strSelect, " B_ITEM A (NOLOCK),    B_ITEM_BY_PLANT B (NOLOCK),  B_STORAGE_LOCATION C (NOLOCK) ", strWhere, lgF0) = False Then
		Call DisplayMsgBox("122700","X", Frm1.vspdData2.Text,"X")
		Call LookUpItemByPlantFail2(Frm1.vspdData2.Text, Row)	    
		Exit Function
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	With frm1.vspdData2
		
		If lgF0(6) = "N"  Or lgF0(7) = "N" Then 'Invalid Item
			Call DisplayMsgBox("122619", "x", "x", "x") 
			Call LookUpItemByPlantFail2(Frm1.vspdData2.Text, Row)
		Else
			If lgF0(5) = "Y" Then
				Call DisplayMsgBox("189214", "x", "x", "x")
			    Call LookUpItemByPlantFail2(FilterVar(Frm1.vspdData2.Text, "''", "S"), Row)
			Else
				.Col = C_ItemNm2
				.text = lgF0(3)
				.Col = C_Spec2
				.text = lgF0(4)
				.Col = C_BaseUnit2
				.text = lgF0(2)

				If lgF0(16) = "N" Then 'TRACKING_FLG
					.Col = C_TrackingNo2
					.Text = "*"
				Else
					.Col = C_TrackingNo2		
					.Value = lgF0(16)
				End If

				.Col = C_SlCd2
				.text = lgF0(12)
				.Col = C_SlNm2
				.text = lgF0(13)
				.Col = C_IssueMthd2
				.text = lgF0(14)
				.Col = C_IssueMthdDesc2
				.text = lgF0(15)    
			End If
		End If
	
	End With
	
	Call LookUpItemByPlantSuccess2(Row)

End Function

Function LookUpItemByPlantFail2(Byval strItemCd, Byval Row)

Dim	strOprNo

    With frm1.vspddata2
		.Row = Row
		.Col = C_ItemCd2
		.text = ""
		.Col = C_ItemNm2
		.text = ""
		.Col = C_Spec2
		.text = ""
		.Col = C_BaseUnit2
		.text = ""
		.Col = C_TrackingNo2
		.text = ""
		.Col = C_SlCd2
		.text = ""
		.Col = C_SlNm2
		.text = ""
		.Col = C_IssueMthd2
		.text = ""
		.Col = C_IssueMthdDesc2
		.text = ""
		.Col = C_OprNo2
		strOprNo = .text
		
	End With
	
	Call SetActiveCell(frm1.vspdData2, C_ItemCd2, Row, "M","X","X")
	Set gActiveElement = document.activeElement
End Function

Function LookUpItemByPlantSuccess2(Byval Row)
	
	Dim strCompntCd
	
	ggoSpread.Source = frm1.vspdData2
	frm1.vspdData2.Row = Row
	frm1.vspdData2.Col = C_ItemCd2
	strCompntCd = frm1.vspdData2.Text

	ggoSpread.UpdateRow Row

'	frm1.vspdData2.Col = C_HndCompntCd
'	frm1.vspdData2.Text = strCompntCd
	
End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd2(Byval strCode, Byval strName, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"											' 팝업 명칭 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE 명칭 
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = strName													' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "창고"												' TextBox 명칭 
   	arrField(0) = "SL_CD"													' Field명(0)
   	arrField(1) = "SL_NM"													' Field명(1)
   	arrHeader(0) = "창고"												' Header명(0)
   	arrHeader(1) = "창고명"												' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSLCd2(arrRet, Row)
	End If
	
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd2(byval arrRet, Byval Row)

    With frm1.vspdData2
	   	.Row = Row
	   	.Col = C_SlCd2
	   	.Text = arrRet(0)
	   	.Col = C_SlNm2
	   	.Text = arrRet(1)
	End With

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

End Function

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If  
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgIntPrevKey <> 0 Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If LayerShowHide(1) = False Then Exit Sub
			If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If	
		End If     
    End if
    
End Sub

'------------------------------------------  OpenItem()  -------------------------------------------------
' Name : OpenBsItem()
' Description : OpenItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBsItem()

	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
			 
	IsOpenPop = True

	arrParam(0) = "기준품목" 
	arrParam(1) = "B_Item a, B_Item b"    
	 
	arrParam(2) = Trim(frm1.txtBsitemcd.Value)
	 
	arrParam(4) = "a.base_item_cd = b.item_cd"   
	arrParam(5) = "기준품목"   
	 
	arrField(0) = "a.base_item_cd" 
	arrField(1) = "b.item_nm" 
	    
	arrHeader(0) = "기준품목"  
	arrHeader(1) = "기준품목명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtBsItemCd.focus
		Exit Function
	Else
		frm1.txtBsitemcd.Value    = arrRet(0)  
		frm1.txtBsitemNm.Value    = arrRet(1)  
		frm1.txtBsitemcd.focus
		Set gActiveElement = document.activeElement
	End If  
End Function
