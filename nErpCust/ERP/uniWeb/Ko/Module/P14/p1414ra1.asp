<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1 %>
<!--======================================================================================================
'*  1. Module Name          : Production																*
'*  2. Function Name        : Reference Popup BOM Copy													*
'*  3. Program ID           : b1b11pa2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : BOM Copy Popup															*
'*  7. Modified date(First) : 2003/03/14																*
'*  8. Modified date(Last)  : 																*
'*  9. Modifier (First)     : Hong Chang Ho																*
'* 10. Modifier (Last)      : 																*
'* 11. Comment              :																			*
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_ID = "p1414rb1.asp"							'☆: Asp name of Biz logic

Const C_SHEETMAXROWS = 30

Dim C_Select
Dim C_Level
Dim C_Seq
Dim C_ChildItemCd
Dim C_ChildItemNm
Dim C_Spec
Dim C_ChildItemUnit
Dim C_ItemAcct
Dim C_ItemAcctNm
Dim C_ProcType
Dim C_ProcTypeNm
Dim C_BomType
Dim C_ChildItemBaseQty
Dim C_ChildBasicUnit
Dim C_PrntItemBaseQty
Dim C_PrntBasicUnit
Dim C_SafetyLT
Dim C_LossRate
Dim C_SupplyFlg
Dim C_SupplyFlgNm
Dim C_ValidFromDt
Dim C_ValidToDt
Dim C_ECNNo
Dim C_ECNDesc
Dim C_ReasonCd
Dim C_ReasonNm
Dim C_DrawingPath
Dim C_Remark
Dim C_HdrItemCd
Dim C_HdrBomNo
Dim C_HdrProcType
Dim C_ItemValidFromDt
Dim C_ItemValidToDt
Dim C_Row

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim strReturn
Dim lgCurDate
Dim gblnWinEvent
Dim IsOpenPop
Dim arrReturn
Dim arrParam					
Dim arrField

Dim lgStrPlantCd
Dim lgStrPrntLevel
Dim lgStrPrntItemCd
Dim lgStrPrntItemAcct
Dim lgStrPrntBOMNo
Dim lgStrPrntProcType
Dim lgStrPrntBasicUnit
Dim lgStrEcnNo
Dim lgStrEcnDesc
Dim lgStrReasonCd
Dim lgStrReasonNm
Dim lgSelectAll

Dim arrParent
			
Dim PopupParent
		
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
arrField = arrParent(2)

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName
	
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Select				= 1
	C_Level					= 2
	C_Seq					= 3
	C_ChildItemCd			= 4
	C_ChildItemNm			= 5
	C_Spec					= 6
	C_ChildItemUnit			= 7
	C_ItemAcct				= 8
	C_ItemAcctNm			= 9
	C_ProcType				= 10
	C_ProcTypeNm			= 11
	C_BomType				= 12
	C_ChildItemBaseQty		= 13
	C_ChildBasicUnit		= 14
	C_PrntItemBaseQty		= 15
	C_PrntBasicUnit			= 16
	C_SafetyLT				= 17
	C_LossRate				= 18
	C_SupplyFlg				= 19
	C_SupplyFlgNm			= 20
	C_ValidFromDt			= 21
	C_ValidToDt				= 22
	C_ECNNo					= 23
	C_ECNDesc				= 24
	C_ReasonCd				= 25
	C_ReasonNm				= 26
	C_DrawingPath			= 27
	C_Remark				= 28
	C_HdrItemCd				= 29
	C_HdrBomNo				= 30
	C_HdrProcType			= 31
	C_ItemValidFromDt		= 32
	C_ItemValidToDt			= 33
	C_Row					= 34
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Function InitVariables()

	lgStrPrevKeyIndex = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	
	lgIntFlgMode = PopupParent.OPMD_CMODE
	gblnWinEvent = False
	
    lgSortKey = 1                                       '⊙: initializes sort direction
	Redim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
	
'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	
	lgStrPlantCd = arrParam(0)
	lgStrPrntLevel = arrParam(2)
	lgStrPrntItemCd = arrParam(3)
	lgStrPrntItemAcct = arrParam(4)
	lgStrPrntBOMNo = arrParam(5)
	lgStrPrntProcType = arrParam(6)
	lgStrPrntBasicUnit = arrParam(7)
	lgStrEcnNo = arrParam(8)
	lgStrEcnDesc = arrParam(9)
	lgStrReasonCd = arrParam(10)
	lgStrReasonNm = arrParam(11)
	
	lgSelectAll = False
	
	txtPlantCd.value = arrParam(0)
	txtPlantNm.value = arrParam(1)
	txtBaseDt.text = StartDate
	txtBomNo.value = "1"
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	    
	Call InitSpreadPosVariables()

    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20030314",, PopupParent.gAllowDragDropSpread

    vspdData.ReDraw = False

    vspdData.MaxCols = C_Row
    vspdData.MaxRows = 0
	    
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetCheck 	C_Select,				"",	2,,, 1
	ggoSpread.SSSetEdit 	C_Level, 				"레벨", 8
	ggoSpread.SSSetFloat	C_Seq,					"순서", 6, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, 1, FALSE, "Z" 
	ggoSpread.SSSetEdit		C_ChildItemCd,			"자품목", 20,,, 18, 2
	ggoSpread.SSSetEdit 	C_ChildItemNm, 			"자품목명", 30
	ggoSpread.SSSetEdit 	C_Spec,	 				"규격", 30
	ggoSpread.SSSetEdit		C_ChildItemUnit,		"단위", 6,,, 3, 2
	ggoSpread.SSSetEdit		C_ItemAcct,				"품목계정", 10
	ggoSpread.SSSetEdit		C_ItemAcctNm,			"품목계정", 10
	ggoSpread.SSSetEdit 	C_ProcType, 			"조달구분", 10
	ggoSpread.SSSetEdit 	C_ProcTypeNm, 			"조달구분", 12
	ggoSpread.SSSetEdit		C_BomType,				"BOM Type", 10,,, 3, 2
	ggoSpread.SSSetFloat	C_ChildItemBaseQty,		"자품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,, "Z"     'hanc
	ggoSpread.SSSetEdit 	C_ChildBasicUnit,		"단위"			, 6,,, 3, 2
	ggoSpread.SSSetFloat	C_PrntItemBaseQty,		"모품목기준수"	, 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec,,, "Z"
	ggoSpread.SSSetEdit		C_PrntBasicUnit,		"단위"			, 6,,, 3, 2
	ggoSpread.SSSetFloat 	C_SafetyLT, 			"안전L/T"	, 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, 1, FALSE, "Z" 
	ggoSpread.SSSetFloat	C_LossRate,				"Loss율"	, 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, 1, FALSE, "Z" 
	ggoSpread.SSSetEdit		C_SupplyFlg,			"유무상구분", 8
	ggoSpread.SSSetEdit		C_SupplyFlgNm,			"유무상구분", 10
	ggoSpread.SSSetDate		C_ValidFromDt,			"시작일"	, 11, 2, PopupParent.gDateFormat
	ggoSpread.SSSetDate 	C_ValidToDt,			"종료일"	, 11, 2, PopupParent.gDateFormat

	ggoSpread.SSSetEdit		C_ECNNo,				"설계변경번호", 18,,, 18, 2
	ggoSpread.SSSetEdit		C_ECNDesc,				"설계변경내용", 30,,, 100
	ggoSpread.SSSetEdit		C_ReasonCd,				"설계변경근거", 10,,, 2, 2
	ggoSpread.SSSetEdit		C_ReasonNm,				"설계변경근거명", 14
	ggoSpread.SSSetEdit		C_DrawingPath,			"도면경로", 30,,, 100

	ggoSpread.SSSetEdit 	C_Remark,	 			"비고"		, 30,,, 1000
	ggoSpread.SSSetEdit		C_HdrItemCd,			"Header품목", 5
	ggoSpread.SSSetEdit		C_HdrBomNo,				"header BOM No.", 5
	ggoSpread.SSSetEdit		C_HdrProcType,			"조달구분", 8
	
	ggoSpread.SSSetDate		C_ItemValidFromDt,		"품목시작일"	, 11, 2, PopupParent.gDateFormat
	ggoSpread.SSSetDate 	C_ItemValidToDt,		"품목종료일"	, 11, 2, PopupParent.gDateFormat
	ggoSpread.SSSetEdit		C_Row,					"순서", 5

	ggoSpread.SSSetSplit2(4)											'frozen 기능 추가 
	
	Call ggoSpread.SSSetColHidden(C_ChildItemUnit, C_ItemAcct, True)
	Call ggoSpread.SSSetColHidden(C_ProcType, C_ProcType, True)
	Call ggoSpread.SSSetColHidden(C_SupplyFlg, C_SupplyFlg, True)
	Call ggoSpread.SSSetColHidden(C_HdrItemCd, C_ItemValidToDt, True)
	Call ggoSpread.SSSetColHidden(C_Row, C_Row, True)
    
	vspdData.ReDraw = True

	Call SetSpreadLock 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = vspdData
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
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Select				= iCurColumnPos(1)
			C_Level					= iCurColumnPos(2)
			C_Seq					= iCurColumnPos(3)
			C_ChildItemCd			= iCurColumnPos(4)
			C_ChildItemNm			= iCurColumnPos(5)
			C_Spec					= iCurColumnPos(6)
			C_ChildItemUnit			= iCurColumnPos(7)
			C_ItemAcct				= iCurColumnPos(8)
			C_ItemAcctNm			= iCurColumnPos(9)
			C_ProcType				= iCurColumnPos(10)
			C_ProcTypeNm			= iCurColumnPos(11)
			C_BomType				= iCurColumnPos(12)
			C_ChildItemBaseQty		= iCurColumnPos(13)
			C_ChildBasicUnit		= iCurColumnPos(14)
			C_PrntItemBaseQty		= iCurColumnPos(15)
			C_PrntBasicUnit			= iCurColumnPos(16)
			C_SafetyLT				= iCurColumnPos(17)
			C_LossRate				= iCurColumnPos(18)
			C_SupplyFlg				= iCurColumnPos(19)
			C_SupplyFlgNm			= iCurColumnPos(20)
			C_ValidFromDt			= iCurColumnPos(21)
			C_ValidToDt				= iCurColumnPos(22)
			C_ECNNo					= iCurColumnPos(23)
			C_ECNDesc				= iCurColumnPos(24)
			C_ReasonCd				= iCurColumnPos(25)
			C_ReasonNm				= iCurColumnPos(26)
			C_DrawingPath			= iCurColumnPos(27)
			C_Remark				= iCurColumnPos(28)
			C_HdrItemCd				= iCurColumnPos(29)
			C_HdrBomNo				= iCurColumnPos(30)
			C_HdrProcType			= iCurColumnPos(31)
			C_ItemValidFromDt		= iCurColumnPos(32)
			C_ItemValidToDt			= iCurColumnPos(33)
			C_Row					= iCurColumnPos(34)
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
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal str)
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(txtPlantCd.value)   ' Plant Code
	arrParam(1) = Trim(str)	' Item Code
	arrParam(2) = ""												' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 

	arrField(0) = 1		'ITEM_CD
    arrField(1) = 2 	'ITEM_NM											
    arrField(2) = 3 	'SPECIFICATION
    arrField(3) = 4 	'BASIC_UNIT
    arrField(4) = 5		'ITEM_ACCT
    arrField(5) = 6		'ITEM_ACCT
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtItemCd.focus
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo(ByVal strItem, ByVal strBom)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
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
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	txtBomNo.focus
	
End Function

Function SetItemCd(ByVal arrRet)
	txtItemCd.Value    = arrRet(0)		
	txtItemNm.Value    = arrRet(1)
	txtSpec.value		= arrRet(2)
	txtBasicUnit.value = arrRet(3)
	txtItemAcct.value = arrRet(4)
	txtItemAcctNm.value = arrRet(5)
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup에서 return된 값 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
	txtBomNo.Value    = arrRet(0)
End Function

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = vspdData
    Call SetPopupMenuItemInf("0000110111")

    If vspdData.MaxRows <= 0 Then Exit Sub

End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'=======================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyDown
'   Event Desc :
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
	If KeyAscii=27 Then
 		Call CancelClick()
	ElseIf KeyAscii = 13 and vspdData.ActiveRow > 0 Then
		Call OkClick()
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				'DbQuery
				If DbQuery = False Then
					Exit Sub
				End If
			End If
		End If
	End With
End Sub
	

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc :
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			'DbQuery
			If DbQuery = False Then
				Exit Sub
			End If
		End If
	End If
End Sub

'========================================================================================================
'	Name : FncQuery
'	Desc : 
'========================================================================================================
Function FncQuery()

	FncQuery = False

    If Not chkField(Document, "1") Then									
       Exit Function
    End If
	
			
	vspdData.MaxRows = 0						'Grid 초기화 
		
	If DbQuery = False Then
		Exit Function
	End If
	
	FncQuery = True
	
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

Function DbQuery()
    'Err.Clear                                                               '☜: Protect system from crashing
	
    DbQuery = False                                                         '⊙: Processing is NG
	
	 '-----------------------
    'Check condition area
    '----------------------- 

	Call LayerShowHide(1)												<%'⊙: 작업진행중 표시 %>	
	    
    Dim strVal
         
  	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001												'☜: 
	strVal = strVal & "&txtPlantCd=" & lgStrPlantCd									'☆: 조회 조건 데이타 
	strVal = strVal & "&txtItemCd=" & Trim(txtItemCd.value)					'☆: 조회 조건 데이타 
	strVal = strVal & "&txtBomNo=" & Trim(txtBomNo.value) 
	strVal = strVal & "&txtBaseDt=" & Trim(txtBaseDt.Text)
	strVal = strval & "&rdoSrchType=2" ' 일단 다단계만 & rdoSrchType1.value 
		
	strVal = strVal & "&txtMaxRows="         & vspdData.MaxRows
    strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
    strVal = strVal & "&lgMaxCount="         & CStr(C_SHEETMAXROWS)    '☜: Max fetched data at a time
        
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
		
    DbQuery = True                                                          '⊙: Processing is NG
    
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk(LngMaxRow)
	Dim iIntCnt, iStrLevel, iIntRow, iStrPrntBOMType, iStrChildItemCd
	
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    Set gActiveElement = document.activeElement
    
    vspdData.ReDraw = False
	
	For iIntCnt = LngMaxRow To vspdData.MaxRows
		Call vspdData.GetText(C_Level, iIntCnt, iStrLevel)
		Call vspdData.GetText(C_ChildItemCd, iIntCnt, iStrChildItemCd)
		
		If Replace(Trim(iStrLevel), ".", "") = "1" Then
			iIntRow = iIntCnt
			ggoSpread.SpreadUnLock	C_Select, iIntCnt, C_Select, iIntCnt
		ElseIf lgStrPrntBOMNo = "1" And Replace(Trim(iStrLevel), ".", "") > 1 Then
			Call vspdData.GetText(C_HdrBomNo, iIntCnt, iStrPrntBOMType)
			If iStrPrntBOMType = "" Then
				ggoSpread.SpreadLock	C_Select, iIntRow, C_Select, iIntRow
			End If
		End If
		If iStrChildItemCd = lgStrPrntItemCd Then
				ggoSpread.SpreadLock	C_Select, iIntRow, C_Select, iIntRow
		End If
	Next					
	
	lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgSelectAll = False
			
	vspdData.ReDraw = True

End Function

Function DbQueryNotOk()
    lgIntFlgMode = PopupParent.OPMD_CMODE								'⊙: Indicates that current mode is Update mode
End Function

'========================================================================================================
'	Name : OKClick
'	Desc : 
'========================================================================================================
Function OKClick()
	Dim iIntCnt, iStrItemCd, iStrItemNm, iStrSpec, iStrItemUnit, iStrItemAcct, iStrItemAcctNm, iStrProcType, istrProcTypeNm
	Dim iStrBomType, iStrChildItemBaseQty, iStrChildBasicUnit, iStrPrntItemBaseQty, iStrSafetyLT, iStrLossRate,iStrDrawingPath,iStrRemark
	Dim iStrSupplyFlg, iStrSupplyFlgNm, iStrChkStatus, iCurColumnPos
	
	If vspdData.MaxRows >= 1 Then
		
        ggoSpread.Source = PopupParent.frbody.document.vspdData
		
		With PopupParent.frbody.document.vspdData
			.Redraw = False
			
			For iIntCnt = 1 To vspdData.MaxRows
				Call vspdData.GetText(1, iIntCnt, iStrChkStatus)
				If iStrChkStatus = "1" Then                        'If Check Box is checked
					Call vspdData.GetText(C_ChildItemCd,	iIntCnt, iStrItemCd)
					Call vspdData.GetText(C_ChildItemNm,	iIntCnt, iStrItemNm)
					Call vspdData.GetText(C_Spec,			iIntCnt, iStrSpec)
					Call vspdData.GetText(C_ChildItemUnit,	iIntCnt, iStrItemUnit)
					Call vspdData.GetText(C_ItemAcct,		iIntCnt, iStrItemAcct)
					Call vspdData.GetText(C_ItemAcctNm,		iIntCnt, iStrItemAcctNm)
					Call vspdData.GetText(C_ProcType,		iIntCnt, iStrProcType)
					Call vspdData.GetText(C_ProcTypeNm,		iIntCnt, iStrProcTypeNm)
					Call vspdData.GetText(C_ChildItemBaseQty,iIntCnt, iStrChildItemBaseQty)
					Call vspdData.GetText(C_ChildBasicUnit,	iIntCnt, iStrChildBasicUnit)
					Call vspdData.GetText(C_PrntItemBaseQty,iIntCnt, iStrPrntItemBaseQty)
					Call vspdData.GetText(C_SafetyLT,		iIntCnt, iStrSafetyLT)
					Call vspdData.GetText(C_LossRate,		iIntCnt, iStrLossRate)
					Call vspdData.GetText(C_SupplyFlg,		iIntCnt, iStrSupplyFlg)
					Call vspdData.GetText(C_SupplyFlgNm,	iIntCnt, iStrSupplyFlgNm)
					Call vspdData.GetText(C_DrawingPath,	iIntCnt, iStrDrawingPath)
					Call vspdData.GetText(C_Remark,			iIntCnt, istrRemark)				
					ggoSpread.InsertRow, 1
					Call .SetText(PopupParent.frbody.C_Level,			.ActiveRow, String(Replace(lgStrPrntLevel, ".", "") + 1, ".") & Replace(lgStrPrntLevel, ".", "") + 1)
					Call .SetText(PopupParent.frbody.C_ChildItemCd,		.ActiveRow, iStrItemCd)
					Call .SetText(PopupParent.frbody.C_ChildItemNm,		.ActiveRow, iStrItemNm)
					Call .SetText(PopupParent.frbody.C_Spec,			.ActiveRow, iStrSpec)
					Call .SetText(PopupParent.frbody.C_ChildItemUnit,	.ActiveRow, iStrItemUnit)
					Call .SetText(PopupParent.frbody.C_ItemAcct,		.ActiveRow, iStrItemAcct)
					Call .SetText(PopupParent.frbody.C_ItemAcctNm,		.ActiveRow, iStrItemAcctNm)
					Call .SetText(PopupParent.frbody.C_ProcType,		.ActiveRow, iStrProcType)
					Call .SetText(PopupParent.frbody.C_ProcTypeNm,		.ActiveRow, iStrProcTypeNm)
					Call .SetText(PopupParent.frbody.C_BomType,			.ActiveRow, lgStrPrntBomNo)
					Call .SetText(PopupParent.frbody.C_ChildItemBaseQty,.ActiveRow, iStrChildItemBaseQty)
					Call .SetText(PopupParent.frbody.C_ChildBasicUnit,	.ActiveRow, iStrChildBasicUnit)
					Call .SetText(PopupParent.frbody.C_PrntItemBaseQty,	.ActiveRow, iStrPrntItemBaseQty)
					Call .SetText(PopupParent.frbody.C_PrntBasicUnit,	.ActiveRow, lgStrPrntBasicUnit)
					Call .SetText(PopupParent.frbody.C_SafetyLT,		.ActiveRow, iStrSafetyLT)
					Call .SetText(PopupParent.frbody.C_LossRate,		.ActiveRow, iStrLossRate)
					Call .SetText(PopupParent.frbody.C_SupplyFlg,		.ActiveRow, iStrSupplyFlg)
					Call .SetText(PopupParent.frbody.C_SupplyFlgNm,		.ActiveRow, iStrSupplyFlgNm)
					Call .SetText(PopupParent.frbody.C_HdrItemCd,		.ActiveRow, lgStrPrntItemCd)
					Call .SetText(PopupParent.frbody.C_HdrBomNo,		.ActiveRow, lgStrPrntBomNo)
					Call .SetText(PopupParent.frbody.C_HdrProcType,		.ActiveRow, lgStrPrntProcType)
					Call .SetText(PopupParent.frbody.C_ValidFromDt,		.ActiveRow, StartDate)
					Call .SetText(PopupParent.frbody.C_ValidToDt,		.ActiveRow, UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"))
					Call .SetText(PopupParent.frbody.C_ECNNo,			.ActiveRow, lgStrEcnNo)
					Call .SetText(PopupParent.frbody.C_ECNDesc,			.ActiveRow, lgStrEcnDesc)
					Call .SetText(PopupParent.frbody.C_ReasonCd,		.ActiveRow, lgStrReasonCd)
					Call .SetText(PopupParent.frbody.C_ReasonNm,		.ActiveRow, lgStrReasonNm)
					Call .SetText(PopupParent.frbody.C_DrawingPath,		.ActiveRow, iStrDrawingPath)
					Call .SetText(PopupParent.frbody.C_Remark,			.ActiveRow, iStrRemark)
					
					Call PopupParent.frbody.SetSpreadColor(.ActiveRow, .ActiveRow, Replace(lgStrPrntLevel, ".", "") + 1, 0)

					If lgStrPrntProcType= "O" Then					'상위품목이 외주가공품인 경우 
						ggoSpread.SpreadUnLock PopupParent.frbody.C_SupplyFlgNm,	.ActiveRow, PopupParent.frbody.C_SupplyFlgNm, .ActiveRow
						ggoSpread.SSSetRequired	PopupParent.frbody.C_SupplyFlgNm,	.ActiveRow, .ActiveRow
					Else
						ggoSpread.SSSetProtected PopupParent.frbody.C_SupplyFlgNm,	.ActiveRow, .ActiveRow
					End If

					If PopupParent.frbody.lgStrBOMHisFlg = "Y" Then
						ggoSpread.SpreadUnLock	PopupParent.frbody.C_ECNNo,		.ActiveRow, PopupParent.frbody.C_ECNNo, .ActiveRow
						ggoSpread.SpreadUnLock	PopupParent.frbody.C_ECNNoPopup,.ActiveRow, PopupParent.frbody.C_ECNNoPopup, .ActiveRow
						ggoSpread.SSSetRequired	PopupParent.frbody.C_ECNNo,		.ActiveRow, .ActiveRow
					Else
						ggoSpread.SSSetProtected PopupParent.frbody.C_ECNNo,	.ActiveRow, .ActiveRow
						ggoSpread.SSSetProtected PopupParent.frbody.C_ECNNoPopup,	.ActiveRow, .ActiveRow
						ggoSpread.SSSetProtected PopupParent.frbody.C_ECNDesc,	.ActiveRow, .ActiveRow
						ggoSpread.SSSetProtected PopupParent.frbody.C_ReasonCd,	.ActiveRow, .ActiveRow
						ggoSpread.SSSetProtected PopupParent.frbody.C_ReasonCdPopup,	.ActiveRow, .ActiveRow
					End If
					
				End If
			Next
			
			.Redraw = True

		End With
	End If
	
	

	Self.Close()
					
End Function

'========================================================================================================
'	Name : CancelClick
'	Desc : 
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
'	Name : MousePointer
'	Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call FncQuery()
	End If
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif", "../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
		
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6", "5", "0")
	Call AppendNumberPlace("7", "2", "2")
	Call AppendNumberPlace("8", "11", "6")
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec, FALSE,, ggStrMinPart, ggStrMaxPart)
    
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call SetDefaultVal()
	Call InitVariables
	Call InitSpreadSheet()

	txtItemCd.focus
	Set gActiveElement = document.activeElement  

End Sub

Function SelectAll()
	Dim iRow	
	
	If vspdData.MaxRows < 1 Then Exit Function
	
	If lgSelectAll = False Then
		with vspdData	
			for iRow=1 to vspdData.MaxRows
				.Row = iRow
				.Col = C_Level 
				If Trim(.Text) = ".1" Then
					.Col = C_Select	
					.text="1"	
				End If			
			next 		
		end with	
	
		lgSelectAll = True
	Else
		with vspdData	
			for iRow=1 to vspdData.MaxRows
				.Row = iRow
				.Col = C_Level
				If Trim(.Text) = ".1" Then
					.Col = C_Select	
					.text="0"	
				End If			
			next 		
		end with	
	
		lgSelectAll = False
	End If	
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/Uni2kCM.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
-->
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14XXXU" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtBaseDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="기준일" tag="12" VIEWASTEXT id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>BOM Type</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBomNo" SIZE=5 MAXLENGTH=3 tag="12XXXU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBomNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBomNo txtItemCd.value, txtBomNo.value"></TD>
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
								<TD CLASS=TD5 NOWRAP>품목계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=15 tag="24" ALT="품목계정"><INPUT TYPE=HIDDEN NAME="txtItemAcct" tag="24" ALT="품목계정"></TD>
								<TD CLASS=TD5 NOWRAP>기준단위</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasicUnit" SIZE=5 MAXLENGTH=3 tag="24" ALT="기준단위"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>규격</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpec" SIZE=40 MAXLENGTH=40 tag="24" ALT="규격"></TD>
								<TD CLASS=TD5 NOWRAP>유효기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemFromDt CLASSID=<%=gCLSIDFPDT%> tag="24" ALT="시작일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemToDt CLASSID=<%=gCLSIDFPDT%> tag="24" ALT="종료일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
								</TD>	
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>BOM 설명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBOMDesc" SIZE=40 tag="24" ALT="BOM 설명">
								<TD CLASS=TD5 NOWRAP>도면경로</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDrawingPath" SIZE=40 MAXLENGTH=40 tag="24" ALT="도면경로">
							</TR>
							<TR>
								<TD HEIGHT="100%" COLSPAN = 4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()"  onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/select_all_d.gif"  Style="CURSOR: hand" ALT="Select All" NAME="Search" ONCLICK="SelectAll()"  onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/select_all.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"	SRC="../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBomType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHdrValidFromDt" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHdrValidToDt" tag="14">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"  TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
