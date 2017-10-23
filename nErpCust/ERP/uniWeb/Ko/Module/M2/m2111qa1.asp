<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name		  : Procurement
'*  2. Function Name		:
'*  3. Program ID		   : M2111QA1
'*  4. Program Name		 : 구매요청조회 
'*  5. Program Desc		 : 구매요청조회 
'*  6. Component List	   :
'*  7. Modified date(First) : 2000/06/08
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)	 : Shin Jin Hyun
'* 10. Modifier (Last)	  : KANG SU HWAN
'* 11. Comment			  :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*							this mark(⊙) Means that "may  change"
'*							this mark(☆) Means that "must change"
'* 13. History			  :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'********************************************************************************************************* !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ============================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ===========================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop
Dim lgSortKey_A
Dim lgPageNo1
Dim lgSortKey_B
Dim lgKeyPos
Dim lgKeyPosVal
Dim	lgTopLeft
Dim IscookieSplit
Dim lgSaveRow
Dim Query_Msg_Flg

Dim C_SpplCd
Dim C_SpplNm
Dim C_QuotaRate
Dim C_ApportionQty
Dim C_PlanDt
Dim C_GrpCd
Dim C_GrpNm
Dim lgPageNo2
Dim lgSpdHdrClicked

Const C_ReqNo 		= 1

Const BIZ_PGM_ID 		= "m2111qb1.asp"
Const BIZ_PGM_ID1	   = "m2111mb1_1.asp"
Const BIZ_PGM_JUMP_ID 	= "m2111ma1"
Const C_MaxKey			  = 19

<%'========================================================================================%>
' Function Name : CookiePage
'========================================================================================
Sub WriteCookiePage()
	Dim strTemp, arrVal

	if frm1.vspdData.ActiveRow > 0 then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",C_ReqNo)
		Call WriteCookie("ReqNo" , frm1.vspdData.Text)
	end if
End Sub

Sub ReadCookiePage()
	If Trim(ReadCookie("m2111ma1_plantcd")) = "" then Exit sub

	frm1.txtPlantCd.Value = ReadCookie("m2111ma1_plantcd")
	frm1.txtItemCd.Value = ReadCookie("m2111ma1_itemcd")

	Call WriteCookie("m2111ma1_plantcd","")
	Call WriteCookie("m2111ma1_itemcd","")

	Call MainQuery()
End Sub


'===================================================================================================================================
Sub InitVariables()

	lgBlnFlgChgValue = False							   'Indicates that no value changed

	lgPageNo   = ""								  'initializes Previous Key for spreadsheet #1
	lgSortKey_A	  = 1
	lgPageNo1   = ""								  'initializes Previous Key for spreadsheet #2
	lgSortKey_B	  = 1

	Query_Msg_Flg		= False
	lgIntFlgMode = parent.OPMD_CMODE
	lgPageNo		 = ""
	lgPageNo1		= ""
	lgPageNo2	= ""
End Sub
'===================================================================================================================================
Sub SetDefaultVal()
 	Dim StartDate, EndDate, EndDate1

	StartDate   = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
	StartDate   = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	EndDate	 = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	EndDate1	= uniDateAdd("m", +1, "<%=GetSvrDate%>", parent.gServerDateFormat)
	EndDate1   = UniConvDateAToB(EndDate1, parent.gServerDateFormat, parent.gDateFormat)

	With frm1
 		.txtDlvyFrDt.Text	= StartDate
 		.txtDlvyToDt.Text	= EndDate1
 		.txtReqFrDt.Text	= StartDate
 		.txtReqToDt.Text	= EndDate
		.txtPlantCd.value	= parent.gPlant
		.txtPlantNm.value	= parent.gPlantNm
		.txtPlantCd.focus
	End With
	Set gActiveElement = document.activeElement
End Sub
'===================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub
<%'===================================================================================================================================%>
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("M2111QA1","S","A","V20040330", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A")

	Call InitSpreadSheet2
End Sub

Sub InitSpreadSheet2()
	Call InitSpreadPosVariables()

	With frm1
		.vspdData2.ReDraw = false

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread

	   .vspdData2.MaxCols = C_GrpNm+1
	   .vspdData2.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_SpplCd, "공급처", 15,,,15,2
		ggoSpread.SSSetEdit	  C_SpplNm, "공급처명", 20
		SetSpreadFloatLocal	  C_QuotaRate, "배분비율(%)",15,1,5
		SetSpreadFloatLocal   C_ApportionQty, "배부량", 15, 1,3
		ggoSpread.SSSetDate	  C_PlanDt, "발주예정일", 15,2,gDateFormat
		ggoSpread.SSSetEdit	  C_GrpCd, "구매그룹", 10,,,10,2
		ggoSpread.SSSetEdit   C_GrpNm, "구매그룹명", 20

		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)

		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols,	.vspdData2.MaxCols,	True)

		.vspdData2.ReDraw = True
	End With

	Call SetSpreadLock("B")
End Sub

'===================================================================================================================================
Sub SetSpreadLock(ByVal pOpt)
	If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

Sub InitSpreadPosVariables()
	C_SpplCd			=	1
	C_SpplNm			=	2
	C_QuotaRate			=	3
	C_ApportionQty		=	4
	C_PlanDt			=	5
	C_GrpCd				=	6
	C_GrpNm				=	7
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
	   Case "A"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpplCd			=	iCurColumnPos(1)
			C_SpplNm			=	iCurColumnPos(2)
			C_QuotaRate			=	iCurColumnPos(3)
			C_ApportionQty		=	iCurColumnPos(4)
			C_PlanDt			=	iCurColumnPos(5)
			C_GrpCd				=	iCurColumnPos(6)
			C_GrpNm				=	iCurColumnPos(7)
	End Select
End Sub

Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
		Case 2															  '금액 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
		Case 3															  '수량 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo	   ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
		Case 4															  '단가 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
		Case 5															  '환율 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
	End Select
End Sub
'===================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_Plant"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "Plant_CD"
	arrField(1) = "Plant_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)
		frm1.txtPlantNm.Value= arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If

End Function
'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)

	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	= arrRet(0)
		frm1.txtItemNm.Value	= arrRet(1)
		frm1.txtItemCd.focus
	End If
End Function

'===================================================================================================================================
Function OpenState()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "요청진행상태"
	arrParam(1) = "B_MINOR"

	arrParam(2) = Trim(frm1.txtStateCd.Value)


	arrParam(4) = "Major_Cd=" & FilterVar("m2101", "''", "S") & ""
	arrParam(5) = "요청진행상태"

	arrField(0) = "Minor_cd"
	arrField(1) = "Minor_Nm"

	arrHeader(0) = "요청진행상태"
	arrHeader(1) = "요청진행상태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtStateCd.focus
		Exit Function
	Else
		frm1.txtStateCd.Value = arrRet(0)
		frm1.txtStateNm.Value = arrRet(1)
		frm1.txtStateCd.focus
	End If
End Function

'===================================================================================================================================
Function OpenDept()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "요청부서"
	arrParam(1) = "B_ACCT_DEPT"

	arrParam(2) = Trim(frm1.txtDeptCd.Value)

	arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(Parent.gChangeOrgId, "''", "S") & " "
	arrParam(5) = "요청부서"

	arrField(0) = "DEPT_CD"
	arrField(1) = "DEPT_NM"

	arrHeader(0) = "요청부서"
	arrHeader(1) = "요청부서명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		frm1.txtDeptCd.Value = arrRet(0)
		frm1.txtDeptNm.Value = arrRet(1)
		frm1.txtDeptCd.focus
	End If
End Function

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim IntRetCD
	Dim iCalledAspName

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
	arrParam(2) = Trim(frm1.txtPlantCd.value)	'공장 
	arrParam(3) = ""	'모품목 
	arrParam(4) = ""	'수주번호 
	arrParam(5) = ""	'추가 Where절 

'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 	iCalledAspName = AskPRAspName("S3135PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackNo.focus
		Exit Function
	Else
		frm1.txtTrackNo.Value = Trim(arrRet)
		frm1.txtTrackNo.focus
	End If
End Function

'==========================================================================================
'   Event Name : OCX_EVENT
'==========================================================================================
 Sub txtDlvyFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDlvyFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDlvyFrDt.Focus
	End if
End Sub

 Sub txtDlvyToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDlvyToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDlvyToDt.Focus
	End if
End Sub

Sub txtReqFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReqFrDt.Focus
	End if
End Sub

Sub txtReqToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtReqToDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtDlvyFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtDlvyToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtReqFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtReqToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


'===================================================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenOrderByPopup(gActiveSpdSheet.Id)
End Sub
'===================================================================================================================================
Function OpenOrderByPopup(ByVal pSpdNo)

	Dim arrRet

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pSpdNo), gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		If Err.Number <> 0 Then
			Err.Clear
		End If
		Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo, arrRet(0), arrRet(1))
	   Call InitVariables
	   Call InitSpreadSheet
   End If
End Function
'===================================================================================================================================
Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")								   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal
	Call AppendNumberPlace("6","5","4")
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")

	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement
End Sub
'===================================================================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)

	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'===================================================================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)

	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub
'===================================================================================================================================
Sub FncSplitColumn()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub
'===================================================================================================================================
Sub vspdData_GotFocus()
	ggoSpread.Source = frm1.vspdData
End Sub
'===================================================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow And NewRow > 0 Then
		Call vspdData_Click(NewCol, NewRow)
		frm1.vspdData2.MaxRows = 0
		Call DbQuery("2", NewRow)
	End If
End Sub
'===================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim ii

	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	SetPopupMenuItemInf("00000000001")

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey_A = 1 Then
			ggoSpread.SSSort, lgSortKey_A
			lgSortKey_A = 2
		Else
			ggoSpread.SSSort, lgSortKey_A
			lgSortKey_A = 1
		End If
		Exit Sub
	End If

	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)

	 lgPageNo1   = ""
	 lgSortKey_B	  = 1
End Sub
'===================================================================================================================================
Sub vspdData2_Click(ByVal Col,  ByVal Row)

	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
	SetPopupMenuItemInf("00000000001")

	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData2
		If lgSortKey_B = 1 Then
			ggoSpread.SSSort, lgSortKey_B
			lgSortKey_B = 2
		Else
			ggoSpread.SSSort, lgSortKey_B
			lgSortKey_B = 1
		End If
		Exit Sub
	End If
End Sub
'===================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)

	If OldLeft <> NewLeft Then
		Exit Sub
	End If


	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then		'☜: 재쿼리 체크 
		If lgPageNo <> "" Then															'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery("1", 0) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub
'===================================================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)

	If OldLeft <> NewLeft Then
		Exit Sub
	End If


	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then		'☜: 재쿼리 체크 
		If lgPageNo1 <> "" Then															'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery("2", frm1.vspdData.ActiveRow) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub
'===================================================================================================================================
Function FncQuery()

	FncQuery = False														'⊙: Processing is NG

	Err.Clear

	Call ggoOper.ClearField(Document, "2")
	Call InitVariables

	If Not chkField(Document, "1") Then
	   Exit Function
	End If

	If DbQuery("1", 0) = False then Exit Function

	FncQuery = True
End Function
'===================================================================================================================================
Function FncPrint()
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'===================================================================================================================================
Function FncExcel()
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'===================================================================================================================================
Function FncFind()
	Call parent.FncFind(parent.C_MULTI , False)
	Set gActiveElement = document.activeElement
End Function
'===================================================================================================================================
Function FncExit()
	FncExit = True
	Set gActiveElement = document.activeElement
End Function
'===================================================================================================================================
Function DbQuery(iOpt, currRow)
	Dim strVal
	Dim strCfmFlg
	DbQuery = False

	If iOpt <> "1" and frm1.vspdData.MaxRows < 1 Then Exit Function

	Err.Clear
	If LayerShowHide(1) = False Then Exit Function

	With frm1

		If iOpt = "1" Then
			If lgIntFlgMode = parent.OPMD_UMODE Then

				strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
				strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
				strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
				strVal = strVal & "&txtItemCd=" & .hdnItem.value
				strVal = strVal & "&txtStateCd=" & .hdnState.value
				strVal = strVal & "&txtDlvyFrDt=" & .hdnDFrDt.Value
				strVal = strVal & "&txtDlvyToDt=" & .hdnDToDt.Value
				strVal = strVal & "&txtReqFrDt=" & .hdnRFrDt.Value
				strVal = strVal & "&txtReqToDt=" & .hdnRToDt.Value
				strVal = strVal & "&txtDeptCd=" & .hdnDept.Value
				strVal = strVal & "&txtTrackNo=" & .hdnTrackNo.Value
				strVal = strVal & "&txtchangorgid=" & parent.gchangeorgid

			Else

				strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
				strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
				strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantcd.value)
				strVal = strVal & "&txtItemCd=" & Trim(.txtItemcd.value)
				strVal = strVal & "&txtStateCd=" & Trim(.txtStateCd.value)
				strVal = strVal & "&txtDlvyFrDt=" & Trim(.txtDlvyFrDt.text)
				strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.text)
				strVal = strVal & "&txtReqFrDt=" & Trim(.txtReqFrDt.text)
				strVal = strVal & "&txtReqToDt=" & Trim(.txtReqToDt.text)
				strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.Value)
				strVal = strVal & "&txtTrackNo=" & Trim(.txtTrackNo.Value)
				strVal = strVal & "&txtchangorgid=" & parent.gchangeorgid
			End if
				strVal = strVal & "&lgPageNo="   & lgPageNo
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="	 & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		Else

			frm1.vspddata.Row = currRow
			frm1.vspdData.Col = GetKeyPos("A",C_ReqNo)

			strVal = BIZ_PGM_ID1 & "?txtPrno=" & Trim(frm1.vspdData.text)
			strVal = strVal & "&lgPageNo="   & lgPageNo1
			strVal = strVal & "&txtMaxRows=" & frm1.vspdData2.MaxRows

		End If

		Call RunMyBizASP(MyBizASP, strVal)
	End With

	DbQuery = True
	Call SetToolbar("1100000000011111")

End Function


'===================================================================================================================================
Function DbQueryOk( iOpt)

  	lgBlnFlgChgValue = False
	lgSaveRow		= 1
	lgIntFlgMode = Parent.OPMD_UMODE
	If iOpt = 1 Then
		If lgTopLeft <> "Y" Then
			Call vspdData_Click(1, 1)
			Call DbQuery("2", 1)
		End If
		lgTopLeft = "N"
		frm1.vspdData.focus
	Else
		Query_Msg_Flg = true
		frm1.vspdData.focus
	End If																 '⊙: This function lock the suitable field
End Function

Function DbQueryOk2()
	Call DbqueryOk(2)
End Function


'===================================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구매요청조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right></td>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant() ">
														   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemCd" SIZE=18 MAXLENGTH=18 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
														   <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
								</TR>
								<tr>
									<TD CLASS="TD5" NOWRAP>요청진행상태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="요청진행상태" NAME="txtStateCd" SIZE=10 MAXLENGTH=5 STYLE="text-transform:uppercase" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenState()">
														   <INPUT TYPE=TEXT ALT="요청진행상태" NAME="txtStateNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
									<TD CLASS="TD5" NOWRAP>필요일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=필요일 NAME="txtDlvyFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=필요일 NAME="txtDlvyToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
									</TD>
								</tr>
								<tr>
									<TD CLASS="TD5" NOWRAP>요청일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=요청일 NAME="txtReqFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=요청일 NAME="txtReqToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
									</TD>
									<TD CLASS="TD5" NOWRAP>요청부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="요청부서" MAXLENGTH=10 NAME="txtDeptCd" SIZE=10  STYLE="text-transform:uppercase" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDept()">
														   <INPUT TYPE=TEXT Alt="요청부서" NAME="txtDeptNm" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
								</tr>
								<tr>
									<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="Tracking No." NAME="txtTrackNo" SIZE=34 MAXLENGTH=25  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingNo()"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</tr>

							</TABLE>
						</FIELDSET>
					</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD HEIGHT=260 WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="A"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>

					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="B"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>

					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>

	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</tr>
	<tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<td WIDTH="*" ALIGN="RIGHT"><a href="VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:WriteCookiePage()">구매요청등록</a></td>
					<td WIDTH="20"></td>
				</tr>
			</table>
		</td>
	</tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="<%=BIZ_PGM_ID%>" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnState" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrsn" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDept" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackNo" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
