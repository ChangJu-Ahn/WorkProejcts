<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b23ma1.asp
'*  4. Program Name         : Manager Item from stye code
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2003/02/04
'*  8. Modified date(Last)  : 2003/07/18
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Ryu Sung Won
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
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID	= "b3b24mb1.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "b3b24mb2.asp"											'☆: 비지니스 로직 ASP명 

Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemFormalNm
Dim C_ItemAcct
Dim C_ItemAcctDesc
Dim C_BasicUnit
Dim C_BasicUnitPopup
Dim C_ItemGroupCd
Dim C_ItemGroupPopup
Dim C_ItemGroupNm
Dim C_Phantom
Dim C_BlanketPurFlg
Dim C_BaseItemCd
Dim C_BaseItemPopup
Dim C_BaseItemNm
Dim C_Spec
Dim C_WeightPerUnit
Dim C_WeightUnit
Dim C_WeightUnitPopup
Dim C_UnitGrossWeight
Dim C_UnitOfGrossWeight
Dim C_GrossUnitPopup
Dim C_CBM
Dim C_CBMDesc
Dim C_DrawingNo
Dim C_HSCd
Dim C_HSCdPopup
Dim C_HSUnit
Dim C_ItemImageFlag
Dim C_VatType
Dim C_VatTypePopup
Dim C_VatTypeDesc
Dim C_VatRate
Dim C_ValidFlag
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_ClassCd
Dim C_ClassNm
Dim C_CharValue1
Dim C_CharValueDesc1
Dim C_CharValue2
Dim C_CharValueDesc2
Dim	C_PlantItemCd

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
Dim StartDate
Dim lgCharCd1
Dim lgCharCd2

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

	C_ItemCd			= 1
	C_ItemNm			= 2
	C_ItemFormalNm		= 3
	C_ItemAcct			= 4
	C_ItemAcctDesc		= 5
	C_BasicUnit			= 6
	C_BasicUnitPopup	= 7
	C_ItemGroupCd		= 8
	C_ItemGroupPopup	= 9
	C_ItemGroupNm		= 10
	C_Phantom			= 11
	C_BlanketPurFlg		= 12
	C_BaseItemCd		= 13
	C_BaseItemPopup		= 14
	C_BaseItemNm		= 15
	C_Spec				= 16
	C_WeightPerUnit		= 17
	C_WeightUnit		= 18
	C_WeightUnitPopup	= 19
	C_UnitGrossWeight	= 20
	C_UnitOfGrossWeight = 21
	C_GrossUnitPopup	= 22
	C_CBM				= 23
	C_CBMDesc			= 24
	C_DrawingNo			= 25
	C_HSCd				= 26
	C_HSCdPopup			= 27
	C_HSUnit			= 28
	C_ItemImageFlag		= 29
	C_VatType			= 30
	C_VatTypePopup		= 31
	C_VatTypeDesc		= 32
	C_VatRate			= 33
	C_ValidFlag			= 34
	C_ValidFromDt		= 35
	C_ValidToDt			= 36
	C_ClassCd			= 37
	C_ClassNm			= 38
	C_CharValue1		= 39
	C_CharValueDesc1	= 40
	C_CharValue2		= 41
	C_CharValueDesc2	= 42
	C_PlantItemCd		= 43

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
    
    lgStrPrevKey = ""
    lgStrPrevKeyIndex = 0                       'initializes Previous Key
    lgStrPrevKeyIndex1 = ""                     'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey    = 1                            'initializes sort direction
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	 'frm1.txtValidFromDt.text  = UniConvYYYYMMDDToDate(parent.gDateFormat, "2003","01","01")
	 frm1.txtValidFromDt.text  = StartDate
	 frm1.txtValidToDt.text	= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
End Sub

'==========================================  2.2.1 SetLocalToolBar()  ========================================
'	Name : SetLocalToolBar()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetLocalToolBar(ByVal pvActiveRow)

	If pvActiveRow < 1 Then Exit Sub

    With frm1

			.vspdData.Col = C_PlantItemCd
			.vspdData.Row = pvActiveRow

			If .vspdData.Text = "" Then
				Call SetToolbar("11001011000111")										'⊙: 버튼 툴바 제어 
			Else
				Call SetToolbar("11001001000111")										'⊙: 버튼 툴바 제어 
			End If
    End With
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
		ggoSpread.Spreadinit "V20030601",, parent.gAllowDragDropSpread

		.ReDraw = False

		.MaxCols = C_PlantItemCd + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_ItemCd,			"품목",20,,,18,2
		ggoSpread.SSSetEdit		C_ItemNm,			"품목명",25,,,40
		ggoSpread.SSSetEdit		C_ItemFormalNm,		"품목정식명칭",30,,,60
		ggoSpread.SSSetCombo	C_ItemAcct,			"품목계정", 10
		ggoSpread.SSSetCombo	C_ItemAcctDesc,		"품목계정", 16
		ggoSpread.SSSetEdit		C_BasicUnit,		"기준단위",10,,,3,2
		ggoSpread.SSSetButton	C_BasicUnitPopup
		ggoSpread.SSSetEdit		C_ItemGroupCd,		"품목그룹",20,,,10,2
		ggoSpread.SSSetButton	C_ItemGroupPopup
		ggoSpread.SSSetEdit		C_ItemGroupNm,		"품목그룹명",20,,,40
		ggoSpread.SSSetCombo	C_Phantom,			"Phantom구분", 12
		ggoSpread.SSSetCombo	C_BlanketPurFlg,	"통합구매구분", 12
		ggoSpread.SSSetEdit		C_BaseItemCd,		"기준품목",20,,,18,2
		ggoSpread.SSSetButton	C_BaseItemPopup
		ggoSpread.SSSetEdit		C_BaseItemNm,		"기준품목명",25,,,40
		ggoSpread.SSSetEdit		C_Spec,				"규격",25,,,50
		ggoSpread.SSSetFloat	C_WeightPerUnit,	"Net중량",15,parent.ggQtyNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"	
		ggoSpread.SSSetEdit		C_WeightUnit,		"Net단위",10,,,3,2
		ggoSpread.SSSetButton	C_WeightUnitPopup
		ggoSpread.SSSetFloat	C_UnitGrossWeight,	 "Gross중량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_UnitOfGrossWeight, "Gross단위",	10,,,3,2
		ggoSpread.SSSetButton 	C_GrossUnitPopup
		ggoSpread.SSSetFloat	C_CBM,				"CBM(부피)",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_CBMDesc,			"CBM정보",	20,,,50	
		ggoSpread.SSSetEdit		C_DrawingNo,		"도면번호",20,,,20
		ggoSpread.SSSetEdit		C_HSCd,				"HS코드",20,,,20,2
		ggoSpread.SSSetButton	C_HSCdPopup
		ggoSpread.SSSetEdit		C_HSUnit,			"HS단위",10,,,3
		ggoSpread.SSSetEdit		C_ItemImageFlag,	"사진유무",10
		ggoSpread.SSSetEdit		C_VatType,			"VAT유형",10,,,5,2
		ggoSpread.SSSetButton	C_VatTypePopup
		ggoSpread.SSSetEdit		C_VatTypeDesc,		"VAT유형명",20
		ggoSpread.SSSetFloat	C_VatRate,			"VAT율",16,parent.ggExchRateNo,ggStrIntegeralPart,ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetCombo	C_ValidFlag,		"유효구분", 10
		ggoSpread.SSSetDate 	C_ValidFromDt,		"시작일", 11, 2, parent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,		"종료일", 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_ClassCd,			"클래스",20,,,16,2
		ggoSpread.SSSetEdit		C_ClassNm,			"클래스명",25
		ggoSpread.SSSetEdit		C_CharValue1,		"사양값1",20,,,16,2
		ggoSpread.SSSetEdit		C_CharValueDesc1,	"사양값명1",25
		ggoSpread.SSSetEdit		C_CharValue2,		"사양값2",20,,,16,2
		ggoSpread.SSSetEdit		C_CharValueDesc2,	"사양값명2",25

		Call ggoSpread.MakePairsColumn(C_BasicUnit, C_BasicUnitPopup )	
		Call ggoSpread.MakePairsColumn(C_ItemGroupCd, C_ItemGroupPopup )	
		Call ggoSpread.MakePairsColumn(C_BaseItemCd, C_BaseItemPopup )
		Call ggoSpread.MakePairsColumn(C_WeightUnit, C_WeightUnitPopup )
		call ggoSpread.MakePairsColumn(C_UnitGrossWeight,	C_GrossUnitPopup)	
		Call ggoSpread.MakePairsColumn(C_HSCd, C_HSCdPopup )	
		Call ggoSpread.MakePairsColumn(C_VatType, C_VatTypePopup )
		
		Call ggoSpread.SSSetColHidden(C_PlantItemCd, C_PlantItemCd, True)
		Call ggoSpread.SSSetColHidden(C_ItemAcct, C_ItemAcct, True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SSSetSplit2(1)										'frozen 기능추가 
	
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
   
		ggoSpread.SpreadLock C_ItemCd, -1, C_ItemCd
		ggoSpread.SpreadLock C_ItemGroupNm, -1, C_ItemGroupNm
		ggoSpread.SpreadLock C_BaseItemNm, -1, C_BaseItemNm
		ggoSpread.SpreadLock C_HSUnit, -1, C_HSUnit
		ggoSpread.SpreadLock C_ItemImageFlag, -1, C_ItemImageFlag
		ggoSpread.SpreadLock C_VatTypeDesc, -1, C_VatTypeDesc
		ggoSpread.SpreadLock C_VatRate, -1, C_VatRate
		ggoSpread.SpreadLock C_ValidFromDt, -1, C_ValidFromDt
		ggoSpread.SpreadLock C_ClassCd, -1, C_ClassCd
		ggoSpread.SpreadLock C_ClassNm, -1, C_ClassNm
		ggoSpread.SpreadLock C_CharValue1, -1, C_CharValue1
		ggoSpread.SpreadLock C_CharValueDesc1, -1, C_CharValueDesc1
		ggoSpread.SpreadLock C_CharValue2, -1, C_CharValue2
		ggoSpread.SpreadLock C_CharValueDesc2, -1, C_CharValueDesc2

		ggoSpread.SSSetProtected .vspdData.MaxCols, -1	
		
		ggoSpread.SSSetRequired  C_ItemNm, -1, C_ItemNm
		ggoSpread.SSSetRequired  C_ItemAcct, -1, C_ItemAcct
		ggoSpread.SSSetRequired  C_ItemAcctDesc, -1, C_ItemAcctDesc
		ggoSpread.SSSetRequired  C_ValidFlag, -1, C_ValidFlag
		ggoSpread.SSSetRequired  C_ValidToDt, -1, C_ValidToDt
		
		.vspdData.ReDraw = True
	
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, byval strFlag)
    With frm1
    
		.vspdData.ReDraw = False

		If strFlag = "Y" Then
			ggoSpread.SSSetRequired 	C_ItemNm, pvStartRow,pvEndRow
		Else
			ggoSpread.SpreadUnLock		C_ItemNm, pvStartRow, C_ItemNm, pvEndRow
		End If

		.vspdData.ReDraw = True
    
    End With
End Sub

'================================== 2.2.5 SetSpreadColor1() ==================================================
' Function Name : SetSpreadColor1
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor1(ByVal pvStartRow)

	Dim LngRow
	Dim LngMaxRow

    With frm1
    
		.vspdData.ReDraw = False
		LngMaxRow = .vspdData.Maxrows 
		
		If pvStartRow = 0 Then pvStartRow = pvStartRow + 1
		
		For LngRow = pvStartRow To LngMaxRow
		
			.vspdData.Col = C_PlantItemCd
			.vspdData.Row = LngRow

			If .vspdData.Text = "" Then
				ggoSpread.SpreadUnLock	 	C_BasicUnit, LngRow,C_BasicUnit
				ggoSpread.SpreadUnLock	 	C_BasicUnitPopup, LngRow,C_BasicUnitPopup
				ggoSpread.SSSetRequired 	C_BasicUnit, LngRow,LngRow
			Else
				ggoSpread.SpreadLock	 	C_BasicUnit, LngRow,C_BasicUnit
				ggoSpread.SpreadLock		C_BasicUnitPopup, LngRow,C_BasicUnitPopup
			End If

		Next

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

			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)
			C_ItemFormalNm		= iCurColumnPos(3)
			C_ItemAcct			= iCurColumnPos(4)
			C_ItemAcctDesc		= iCurColumnPos(5)
			C_BasicUnit			= iCurColumnPos(6)
			C_BasicUnitPopup	= iCurColumnPos(7)
			C_ItemGroupCd		= iCurColumnPos(8)
			C_ItemGroupPopup	= iCurColumnPos(9)
			C_ItemGroupNm		= iCurColumnPos(10)
			C_Phantom			= iCurColumnPos(11)
			C_BlanketPurFlg		= iCurColumnPos(12)
			C_BaseItemCd		= iCurColumnPos(13)
			C_BaseItemPopup		= iCurColumnPos(14)
			C_BaseItemNm		= iCurColumnPos(15)
			C_Spec				= iCurColumnPos(16)
			C_WeightPerUnit		= iCurColumnPos(17)
			C_WeightUnit		= iCurColumnPos(18)
			C_WeightUnitPopup	= iCurColumnPos(19)
			C_UnitGrossWeight	= iCurColumnPos(20) 
			C_UnitOfGrossWeight	= iCurColumnPos(21)
			C_GrossUnitPopup	= iCurColumnPos(22)
			C_CBM				= iCurColumnPos(23) 
			C_CBMDesc			= iCurColumnPos(24)
			C_DrawingNo			= iCurColumnPos(25)
			C_HSCd				= iCurColumnPos(26)
			C_HSCdPopup			= iCurColumnPos(27)
			C_HSUnit			= iCurColumnPos(28)
			C_ItemImageFlag		= iCurColumnPos(29)
			C_VatType			= iCurColumnPos(30)
			C_VatTypePopup		= iCurColumnPos(31)
			C_VatTypeDesc		= iCurColumnPos(32)
			C_VatRate			= iCurColumnPos(33)
			C_ValidFlag			= iCurColumnPos(34)
			C_ValidFromDt		= iCurColumnPos(35)
			C_ValidToDt			= iCurColumnPos(36)
			C_ClassCd			= iCurColumnPos(37)
			C_ClassNm			= iCurColumnPos(38)
			C_CharValue1		= iCurColumnPos(39)
			C_CharValueDesc1	= iCurColumnPos(40)
			C_CharValue2		= iCurColumnPos(41)
			C_CharValueDesc2	= iCurColumnPos(42)
			C_PlantItemCd		= iCurColumnPos(43)

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
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData(1)
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    On Error Resume Next
    Err.Clear

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboItemAcct, "" & Chr(11) & lgF0, "" & Chr(11) & lgF1, Chr(11))
		  
End Sub

'============================= 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'========================================================================================= 
 Sub InitSpreadComboBox()
 
    Dim strCbo  
    
	strCbo = ""    
	strCbo = strCbo & "Y" & vbTab & "N" 
    
	ggoSpread.SetCombo strCbo,C_Phantom
	ggoSpread.SetCombo strCbo,C_BlanketPurFlg
	ggoSpread.SetCombo strCbo,C_ValidFlag
	
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ItemAcct
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ItemAcctDesc
	
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_ItemAcct
			intIndex = .value
			.Col = C_ItemAcctDesc
			.value = intindex
		Next	
	End With
End Sub

'==========================================  2.2.6 InitDataRange()  ========================================== 
'	Name : InitDataRange()
'	Description : Combo Display
'======================================================================================================== 
Sub InitDataRange(ByVal lngStartRow, ByVal lngEndRow)
	Dim intRow
	Dim intIndex

	With frm1.vspdData
		For intRow = lngStartRow To lngEndRow
			.Row = intRow
			.col = C_ItemAcct
			intIndex = .value
			.Col = C_ItemAcctDesc
			.value = intindex
		Next	
	End With
End Sub

'========================================================================================
' Function Name : LookupChar12
' Function Desc : Lookup Characteristic 1/2
'========================================================================================
Function LookupChar12() 

	If gLookUpEnable = False Then Exit Function

    LookupChar12 = False
    
    LayerShowHide(1) 
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & ""				'☜: 
		strVal = strVal & "&txtClassCd=" & Trim(.txtClassCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd1=" & Trim(.txtCharValueCd1.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd2=" & Trim(.txtCharValueCd2.value)	'☆: 조회 조건 데이타 
    
		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    LookupChar12 = True

End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtitemcd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(1) = ""							' Item Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
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

'------------------------------------------  OpenClassCd()  -------------------------------------------------
'	Name : OpenClassCd()
'	Description : Class PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenClassCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtClasscd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtClassCd.value)	' Class Code
	arrParam(1) = ""							' Class Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "Class_CD"
    arrField(1) = 2 							' Field명(1) : "Class_NM"
	
	iCalledAspName = AskPRAspName("B3B31PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B31PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetClassCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtClassCd.focus
	
End Function

'==========================================  OpenCharValueCd()  ==========================================
'	Name : OpenCharValueCd()
'	Description : Open Popup
'========================================================================================================= 
Function OpenCharValueCd(Byval iCallFlag)

	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName, IntRetCD
	Dim strCharValue
	Dim strCharCd
	
	If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
        Exit Function
	End If

	If iCallFlag = 1 Then
		If IsOpenPop = True Or UCase(frm1.txtCharValueCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	Else
		If IsOpenPop = True Or UCase(frm1.txtCharValueCd2.className) = UCase(parent.UCN_PROTECTED) Then Exit Function	
	End If

	IsOpenPop = True

	If iCallFlag = 1 Then
		strCharValue = lgCharCd1
		strCharCd = Trim(frm1.txtCharValueCd1.value)
	Else
		strCharValue = lgCharCd2
		strCharCd = Trim(frm1.txtCharValueCd2.value)
	End If

	arrParam(0) = UCase(Trim(strCharValue))					' CharValue Code
	arrParam(1) = strCharCd
	arrParam(2) = ""										' ----------
	arrParam(3) = ""										' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 										' Field명(0) : "Characteristic_CD"
    arrField(1) = 2 										' Field명(1) : "Characteristic_NM"

	iCalledAspName = AskPRAspName("B3B32PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B32PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=490px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetCharValueCd(iCallFlag, arrRet)
	End If
	
	Call SetFocusToDocument("M")
	If iCallFlag = 1 Then
		frm1.txtCharValueCd1.focus
	Else
		frm1.txtCharValueCd2.focus
	End If
	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit(byval strUnit, byval Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(strUnit)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION <> " & FilterVar("TM", "''", "S") & "  "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet, Row)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_BasicUnit,Row,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenItemGroup()  --------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  AND VALID_TO_DT >=  " & FilterVar("<%=BaseDate%>" , "''", "S") & ""
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

'------------------------------------------  OpenItemGroup1()  --------------------------------------------
'	Name : OpenItemGroup1()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup1(byval strItemGroup, byval Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(strItemGroup)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  AND LEAF_FLG = " & FilterVar("Y", "''", "S") & "  AND VALID_TO_DT >=  " & FilterVar("<%=BaseDate%>" , "''", "S") & ""
	arrParam(5) = "품목그룹"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"
    
    arrHeader(0) = "품목그룹"		
    arrHeader(1) = "품목그룹명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup1(arrRet, Row)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_ItemGroupCd,Row,"M","X","X")
	Set gActiveElement = document.activeElement
	
	
End Function

'------------------------------------------  OpenBasisItemCd()  ------------------------------------------
'	Name : OpenBasisItemCd()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBasisItemCd(byval strBaseItem, byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(strBaseItem)					' Plant Code
	arrParam(1) = ""								' Item Code
	arrParam(2) = ""								' ----------
	arrParam(3) = ""								' ----------
	arrParam(4) = ""

    arrField(0) = 1 								' Field명(0) : "ITEM_CD"
    arrField(1) = 2 								' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBasisItemCd(arrRet, Row)
	End If
	
	Call SetActiveCell(frm1.vspdData,C_BaseItemCd,Row,"M","X","X")
	Set gActiveElement = document.activeElement
		
End Function

'------------------------------------------  OpenWeightUnit()  -------------------------------------------
'	Name : OpenWeightUnit()
'	Description : WeightUnit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWeightUnit(byval strWeightUnit, byval Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(strWeightUnit)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & " "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetWeightUnit(arrRet, Row)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_WeightUnit,Row,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenGrossUnit()  -------------------------------------------
'	Name : OpenGrossUnit()
'	Description : WeightUnit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGrossUnit(byval strWeightUnit, byval Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(strWeightUnit)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION=" & FilterVar("WT", "''", "S") & " "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetGrossUnit(arrRet, Row)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_UnitOfGrossWeight,Row,"M","X","X")
	Set gActiveElement = document.activeElement
	
End Function

'------------------------------------------  OpenHsCd()  -------------------------------------------------
'	Name : OpenHsCd()
'	Description : HS Cd PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenHsCd(byval strHSCd, byval Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HS팝업"	
	arrParam(1) = "B_HS_CODE"				
	arrParam(2) = Trim(strHSCd)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "HS코드"
	
    arrField(0) = "HS_CD"	
    arrField(1) = "HS_NM"
    arrField(2) = "HS_SPEC"	
    arrField(3) = "HS_UNIT"
    	
    
    arrHeader(0) = "HS코드"		
    arrHeader(1) = "HS명"
    arrHeader(2) = "HS규격"
    arrHeader(3) = "HS단위"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetHSCd(arrRet, Row)
	End If
	
	Call SetActiveCell(frm1.vspdData,C_HSCd,Row,"M","X","X")
	Set gActiveElement = document.activeElement	
	
End Function

'===========================================================================
' Function Name : OpenVATType
' Function Desc : OpenVATType Reference Popup
'===========================================================================
Function OpenVATType(byval strVATType, byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(1) = "B_MINOR ,B_CONFIGURATION "	' TABLE 명칭 
	arrParam(2) = Trim(strVATType)				' Code Condition
	arrParam(3) = ""							' Name Condition
	arrParam(4) = "B_MINOR.MAJOR_CD = " & FilterVar("B9001", "''", "S") & " " _
					& " AND B_MINOR.MINOR_CD=B_CONFIGURATION.MINOR_CD " _
					& " AND B_MINOR.MAJOR_CD=B_CONFIGURATION.MAJOR_CD "	_
					& " AND B_CONFIGURATION.SEQ_NO = 1 "					' Where Condition
	arrParam(5) = "VAT유형"					' TextBox 명칭 
		
	arrField(0) = "B_MINOR.MINOR_CD"			' Field명(0)
	arrField(1) = "B_MINOR.MINOR_NM"			' Field명(1)
	arrField(2) = "F5" & parent.gColSep & "B_CONFIGURATION.REFERENCE"				' Field명(2)
	    	    
	arrHeader(0) = "VAT유형"				' Header명(0)
	arrHeader(1) = "VAT유형명"				' Header명(1)
	arrHeader(2) = "VAT율"					' Header명(2)

	arrParam(0) = arrParam(5)					' 팝업 명칭 

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetVATType(arrRet, Row)
	End If	
	
	Call SetActiveCell(frm1.vspdData,C_VatType,Row,"M","X","X")
	Set gActiveElement = document.activeElement	
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetItemCd()  ------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement 		
End Function
'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetClassCd(byval arrRet)
	frm1.txtClassCd.Value    = arrRet(0)		
	frm1.txtClassNm.Value    = arrRet(1)
	
	Call LookUpChar12()
	
	frm1.txtClassCd.focus
	Set gActiveElement = document.activeElement 		
End Function

'==========================================  SetCharValueCd()  ===========================================
'	Name : SetCharValueCd()
'	Description : Set Popup Values
'========================================================================================================= 
Function SetCharValueCd(byval iCallFlag, byval arrRet)

	If iCallFlag = 1 Then
		frm1.txtCharValueCd1.Value	= arrRet(0)	
		frm1.txtCharValueNm1.Value  = arrRet(1)
	Else
		frm1.txtCharValueCd2.Value  = arrRet(0)
		frm1.txtCharValueNm2.Value  = arrRet(1)
	End If
	
	frm1.txtCharValueCd1.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetUnit(byval arrRet, byval Row)

	If Row < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = C_BasicUnit
		.Text = arrRet(0)
	End With
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
End Function

'------------------------------------------  SetItemGroup()  ---------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value	= arrRet(0)		
	frm1.txtItemGroupNm.value   = arrRet(1)
End Function

'------------------------------------------  SetItemGroup1()  ---------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup1(byval arrRet, byval Row)

	If Row < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = C_ItemGroupCd
		.Text = UCase(arrRet(0))
		.Col = C_ItemGroupNm
		.Text = arrRet(1)
	End With

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

End Function

'------------------------------------------  SetBasisItemCd()  -------------------------------------------
'	Name : SetBasisItemCd()
'	Description : BasisItemCd Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetBasisItemCd(byval arrRet, byval Row)
	
	If Row < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = C_BaseItemCd
		.Text = UCase(arrRet(0))
		.Col = C_BaseItemNm
		.Text = arrRet(1)
	End With
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
End Function

'------------------------------------------  SetWeightUnit()  --------------------------------------------
'	Name : SetWeightUnit()
'	Description : WeightUnit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWeightUnit(byval arrRet, byval Row)

	If Row < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = C_WeightUnit
		.Text = arrRet(0)
	End With
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
End Function

'------------------------------------------  SetGrossUnit()  --------------------------------------------
'	Name : SetGrossUnit()
'	Description : WeightUnit Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetGrossUnit(byval arrRet, byval Row)

	If Row < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = C_UnitOfGrossWeight
		.Text = arrRet(0)
	End With
	
	Call vspdData_Change(C_UnitOfGrossWeight, Row)
End Function

'------------------------------------------  SetHSCd()  --------------------------------------------------
'	Name : SetHSCd()
'	Description : HSCd Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetHSCd(byval arrRet, byval Row)

	If Row < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = C_HSCd
		.Text = arrRet(0)
		.Col = C_HSUnit
		.Text = arrRet(3)
	End With
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
End Function

'------------------------------------------  SetVATType()  -----------------------------------------------
'	Name : SetVATType()
'	Description : Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetVATType(byval arrRet, byval Row)

	If Row < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = C_VatType
		.Text = arrRet(0)
		.Col = C_VatTypeDesc
		.Text = arrRet(1)
		.Col = C_VatRate
		.Value = arrRet(2)
	End With

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

End Function

'===========================================  2.5.2 ChkBaseItem(strData1, strData2)  ====================
'=	Event Name : ChkBaseItem(strData1, strData2)
'=	Event Desc : 기준품목과 품목 동일 여부 체크 
'========================================================================================================

Function ChkBaseItem(strData1, strData2)
	
	ChkBaseItem = False
	
	If UCase(Trim(strData1)) = UCase(Trim(strData2)) Then
		Call DisplayMsgBox("127421", "X", "기준품목", "품목")
		
		frm1.txtBasisItemCd.value = ""
		frm1.txtBasisItemNm.value = "" 
		frm1.txtBasisItemCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	ChkBaseItem = True
	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
   
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '⊙: Initializes local global variables
   
    Call InitComboBox
   	Call InitSpreadComboBox
    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
    
	frm1.txtItemCd.focus 
	Set gActiveElement = document.activeElement 
  
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

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

End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtValidFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 MainQuery한다.
'=======================================================================================================
Sub txtValidToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtClassCd_onChange()
'   Event Desc : 
'=======================================================================================================
Sub txtClassCd_onChange()
	If frm1.txtClassCd.value = "" Then
		 frm1.txtClassNm.value = ""
 		 lgCharCd1 = ""
		 lgCharCd2 = ""
	Else
		Call LookUpChar12()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspddata_Click(ByVal Col , ByVal Row )
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000110111")
	Else 	
		Call SetPopupMenuItemInf("0001111111") 
	End If
    
    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	Call SetLocalToolBar(Row)
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
'   Event Name : vspddata_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspddata_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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
			Case  C_ItemAcctDesc
				.Col = Col
				intIndex = .Value
				.Col = C_ItemAcct
				.Value = intIndex
		End Select
    
    End With

End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
		
	End With

End Sub


'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	If Row <= 0 Then Exit Sub

		With frm1.vspdData

		ggoSpread.Source = frm1.vspdData

		.Row = Row

		Select Case Col

		    Case C_BasicUnitPopup

				.Col = C_BasicUnit
				Call OpenUnit(.Text, Row)
				
		    Case C_ItemGroupPopup

				.Col = C_ItemGroupCd
				Call OpenItemGroup1(.Text, Row)
				
		    Case C_BaseItemPopup

				.Col = C_BaseItemCd
				Call OpenBasisItemCd(.Text, Row)

		    Case C_WeightUnitPopup

				.Col = C_WeightUnit
				Call OpenWeightUnit(.Text, Row)
			
			Case C_GrossUnitPopup

				.Col = C_UnitOfGrossWeight
				Call OpenGrossUnit(.Text, Row)
			
		    Case C_HSCdPopup
		    
				.Col = C_HSCd
   				Call OpenHsCd(.Text, Row)

		    Case C_VatTypePopup

				.Col = C_VatType
				Call OpenVATType(.Text, Row)

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
	'----------  Coding part  -------------------------------------------------------------   
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	'----------  Coding part  -------------------------------------------------------------   
	Call SetLocalToolBar(NewRow)

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
		If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
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

    FncQuery = False															'⊙: Processing is NG

    Err.Clear																	'☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtClassCd.value = "" Then
		frm1.txtClassNm.value = ""
	End If
		
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call ggoSpread.ClearSpreadData
'    Call SetDefaultVal
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
    End If     													'☜: Query db data

    FncQuery = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 

    Err.Clear																	'☜: Protect system from crashing
    On Error Resume Next														'☜: Protect system from crashing

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    
    Err.Clear																	'☜: Protect system from crashing
    On Error Resume Next														'☜: Protect system from crashing
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False																'⊙: Processing is NG
    
    Err.Clear																	'☜: Protect system from crashing
    'On Error Resume Next														'☜: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")								'⊙: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then
		Exit Function
	End If
	
    If Not chkField(Document, "3") Then
		Exit Function
	End If
	
	ggoSpread.Source = frm1.vspdData
	If Not ggoSpread.SSDefaultCheck  Then                                  '⊙: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		LayerShowHide(0)
		Exit Function           
    End If     																	'☜: Save db data
    
    FncSave = True																'⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 

	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo						

	Call InitDataRange(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
   
    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
	 
    With frm1.vspdData 
    
    .focus
    Set gActiveElement = document.activeElement 
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow
    
    End With
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
    Call parent.FncFind(parent.C_MULTI, False)	                   '☜:화면 유형, Tab 유무 
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
		
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'☜: 
		
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)					'☆: 조회 조건 데이타 
		strVal = strVal & "&cboItemAcct=" & Trim(.hItemAcct.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&rdoValidFlg=" & Trim(.hrdoValidFlg.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtValidFromDt=" & Trim(.hValidFromDt.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtValidToDt=" & Trim(.hValidToDt.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtClassCd=" & Trim(.hClassCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd1=" & Trim(.hCharValueCd1.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd2=" & Trim(.hCharValueCd2.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'☜: 
		
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&cboItemAcct=" & Trim(.cboItemAcct.value)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)		'☆: 조회 조건 데이타 
		If .rdoValidFlg1.checked = True Then
			strVal = strVal & "&rdoValidFlg=" & ""
		ElseIf .rdoValidFlg2.checked = True Then
			strVal = strVal & "&rdoValidFlg=" & "Y"
		Else
			strVal = strVal & "&rdoValidFlg=" & "N"
		End If
		strVal = strVal & "&txtValidFromDt=" & Trim(.txtValidFromDt.text)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtValidToDt=" & Trim(.txtValidToDt.text)			'☆: 조회 조건 데이타 
		strVal = strVal & "&txtClassCd=" & Trim(.txtClassCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd1=" & Trim(.txtCharValueCd1.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtCharValueCd2=" & Trim(.txtCharValueCd2.value)	'☆: 조회 조건 데이타 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal & "&lgStrPrevKeyIndex1=" & lgStrPrevKeyIndex1
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
	Call InitData(LngMaxRow)

    lgBlnFlgChgValue = False
    
    
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    Call SetLocalToolBar(1)
	
	Call SetSpreadColor1(LngMaxRow)
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
	lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
		
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
on error resume next
err.Clear
    Dim lRow  
	Dim strVal
	Dim strDel
	Dim TmpBufferVal, TmpBufferDel
	Dim iValCnt, iDelCnt
	Dim iTotalStrVal, iTotalStrDel
	Dim iColSep,iRowSep
	
    DbSave = False                                                          '⊙: Processing is NG
    
    '-----------------------
    'Check Valid Date
    '-----------------------
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function  

	LayerShowHide(1) 
		
	With frm1

		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

    iColSep = Parent.gColSep
    iRowSep = parent.gRowSep

    iValCnt = 0
    iDelCnt = 0
    
    ReDim TmpBufferVal(0)
    ReDim TmpBufferDel(0)

    '-----------------------							  
    'Data manipulate area								  
    '-----------------------							  
    For lRow = 1 To .vspdData.MaxRows					  

        Select Case GetSpreadText(.vspdData,0,lRow,"X","X")

            Case ggoSpread.UpdateFlag

				strVal = ""
				strVal = strVal & "U" & iColSep & lRow & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X"))) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemNm,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemFormalNm,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Phantom,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemAcct,lRow,"X","X")) & iColSep
                strVal = strVal & "" & iColSep		' Item Class
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Spec,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_HSCd,lRow,"X","X"))) & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_HSUnit,lRow,"X","X"))) & iColSep                
                strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_WeightPerUnit,lRow,"X","X")),0) & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_WeightUnit,lRow,"X","X"))) & iColSep
                strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_UnitGrossWeight,lRow,"X","X")),0) & iColSep		'2003-07-18
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_UnitOfGrossWeight,lRow,"X","X")) & iColSep
                strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_CBM,lRow,"X","X")),0) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_CBMDesc,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_BasicUnit,lRow,"X","X"))) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_DrawingNo,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_BlanketPurFlg,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_BaseItemCd,lRow,"X","X"))) & iColSep
                strVal = strVal & "0" & iColSep		' Proportion Rate
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_ValidFlag,lRow,"X","X")) & iColSep
                strVal = strVal & UNIConvDate(Trim(GetSpreadText(.vspdData,C_ValidFromDt,lRow,"X","X"))) & iColSep
                strVal = strVal & UNIConvDate(Trim(GetSpreadText(.vspdData,C_ValidToDt,lRow,"X","X"))) & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_VatType,lRow,"X","X"))) & iColSep
                strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_VatRate,lRow,"X","X")),0) & iColSep
				strVal = strVal & "Y" & iColSep		' Class Flag
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_ClassCd,lRow,"X","X"))) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_CharValue1,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_CharValue2,lRow,"X","X")) & iColSep
                strVal = strVal & Trim(UCase(GetSpreadText(.vspdData,C_ItemGroupCd,lRow,"X","X"))) & iRowSep

                ReDim Preserve TmpBufferVal(iValCnt)
                TmpBufferVal(iValCnt) = strVal
                iValCnt = iValCnt + 1
              
            Case ggoSpread.DeleteFlag

				strDel = ""
				strDel = strDel & "D" & iColSep & lRow & iColSep
                strDel = strDel & Trim(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X")) & iRowSep
                
                ReDim Preserve TmpBufferDel(iDelCnt)
                TmpBufferDel(iDelCnt) = strDel
                iDelCnt = iDelCnt + 1

        End Select

    Next

    iTotalStrVal = Join(TmpBufferVal, "")
    iTotalStrDel = Join(TmpBufferDel, "")

	.txtSpread.value = iTotalStrDel & iTotalStrVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
    
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
	
End Function

Function DbDeleteOk()

End Function

'==============================================================================
' Function : ChkValidData
' Description : Start Day와 End Day Check
'==============================================================================
Function ChkValidData(SDay, STime, EDay, ETime)
	ChkValidData = 0

	If CInt(SDay) > CInt(EDay) Then
		ChkValidData = 1
		Exit Function
	End If
	
	If Len(Trim(STime)) <> 8 and Len(Trim(STime)) <> 0 Then
		ChkValidData = -1
		Exit Function
	End IF
	
	If Len(Trim(ETime)) <> 8 and Len(Trim(ETime)) <> 0 Then
		ChkValidData = -2
		Exit Function
	End IF
	
	If CInt(SDay) = CInt(EDay) Then
		If STime > ETime Then
			ChkValidData = 2
			Exit Function
		End If	
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

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목정보관리</font></td>
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
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14" ALT="품목명"></TD>
									<TD CLASS=TD5 NOWRAP>클래스</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClassCd" SIZE=18 MAXLENGTH=16 tag="11XXXU"  ALT="클래스"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClassCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenClassCd()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtClassNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사양값1</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCharValueCd1" SIZE=18 MAXLENGTH=16 tag="11XXXU" ALT="사양값1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharValue1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharValueCd(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtCharValueNm1" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>사양값2</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCharValueCd2" SIZE=18 MAXLENGTH=16 tag="11XXXU" ALT="사양값2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharValue2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharValueCd(2)">&nbsp;<INPUT TYPE=TEXT NAME="txtCharValueNm2" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=18 MAXLENGTH=10 tag="11XXXU"  ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=20 tag="14" ALT="품목그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>품목계정</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="품목계정" STYLE="Width: 150px;" tag="11"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>유효구분</TD>
									<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg1" Value="A" CLASS="RADIO" tag="1X" CHECKED><LABEL FOR="rdoValidFlg1">전체</LABEL>
												<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg2" Value="Y" CLASS="RADIO" tag="1X"><LABEL FOR="rdoValidFlg2">예</LABEL>
												<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg3" Value="N" CLASS="RADIO" tag="1X"><LABEL FOR="rdoValidFlg3">아니오</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>종료일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/b3b24ma1_I674339699_txtValidFromDt.js'></script> &nbsp;~&nbsp;
										<script language =javascript src='./js/b3b24ma1_I490513700_txtValidToDt.js'></script></TD>
								<TR>									
								
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
									<script language =javascript src='./js/b3b24ma1_I861101806_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hClassCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCharValueCd1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCharValueCd2" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hrdoValidFlg" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hValidFromDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hValidToDt" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
