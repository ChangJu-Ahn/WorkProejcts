
Option Explicit

Dim lgIsOpenPop
Dim lgSaveRow

Const BIZ_PGM_ID 		= "m3112qb4_KO441.asp"
Const BIZ_PGM_JUMP_ID 	= ""   '점프없음 
Const C_MaxKey		  = 21

'================================================================================================================================
Sub InitVariables()
	lgPageNo		 = ""
	lgBlnFlgChgValue = False
	lgSortKey		= 1
	lgSaveRow		= 0
	lgIntFlgMode = Parent.OPMD_CMODE
End Sub
'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtDvFrDt.Text	= StartDate
	frm1.txtDvToDt.Text	= EndDate
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd, "Q") 
		frm1.txtPurGrpCd.Tag = left(frm1.txtPurGrpCd.Tag,1) & "4" & mid(frm1.txtPurGrpCd.Tag,3,len(frm1.txtPurGrpCd.Tag))
        frm1.txtPurGrpCd.value = lgPGCd
	End If
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3112QA4","S","A","V20040602", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A")
End Sub
'================================================================================================================================
Sub SetSpreadLock(ByVal pOpt)
	If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub
'================================================================================================================================
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function
    
	lgIsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "PLANT_CD"
	arrField(1) = "PLANT_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function
'================================================================================================================================
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
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
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function
'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "거래처"
	arrParam(1) = "B_Biz_Partner"
	arrParam(2) = Trim(frm1.txtBpCd.Value)
'	arrParam(3) = Trim(frm1.txtBpNm.Value)
	arrParam(4) = "BP_TYPE in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "
	arrParam(5) = "거래처"

	arrField(0) = "BP_CD"
	arrField(1) = "BP_NM"

	arrHeader(0) = "거래처"
	arrHeader(1) = "거래처명"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function
'================================================================================================================================
Function OpenPurGrpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "구매그룹"
	arrParam(1) = "B_Pur_Grp"

	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)

	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "구매그룹"

	arrField(0) = "PUR_GRP"
	arrField(1) = "PUR_GRP_NM"

	arrHeader(0) = "구매그룹"
	arrHeader(1) = "구매그룹명"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus
		Set gActiveElement = document.activeElement
	End If

End Function
'================================================================================================================================
Function OpenPoType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "발주형태"
	arrParam(1) = "M_CONFIG_PROCESS"
	arrParam(2) = Trim(frm1.txtPoType.Value)
'	arrParam(3) = Trim(frm1.txtPoTypeNm.Value)
	arrParam(4) = ""
	arrParam(5) = "발주형태"

	arrField(0) = "PO_TYPE_CD"
	arrField(1) = "PO_TYPE_NM"

	arrHeader(0) = "발주형태"
	arrHeader(1) = "발주형태명"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoType.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoType.Value = arrRet(0)
		frm1.txtPoTypeNm.Value = arrRet(1)
		frm1.txtPoType.focus
		Set gActiveElement = document.activeElement
	End If
End Function
'================================================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenSortPopup("A")
End Sub
'================================================================================================================================
Function OpenSortPopup(ByVal pSpdNo)
	Dim arrRet

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
	   Call InitVariables
	   Call InitSpreadSheet()
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


'================================================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877

	If Kubun = 1 Then

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie CookieSplit , ""
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

'		If arrVal(0) = "" Then Exit Function

		Dim iniSep

		If Len(ReadCookie ("ItemCd")) Then
			frm1.txtItemCd.Value	=  ReadCookie ("ItemCd")
			WriteCookie "ItemCd",""
		Else
			frm1.txtItemCd.Value	=  arrVal(0)
		End If

		frm1.txtItemNm.Value	=  arrVal(1)

		If Len(ReadCookie ("PlantCd")) Then
			frm1.txtPlantCd.Value	=  ReadCookie ("PlantCd")
			WriteCookie "PlantCd",""
		Else
			frm1.txtPlantCd.Value	=  arrVal(3)
		End If

		frm1.txtPlantNm.value	=  arrVal(4)



		If Len(ReadCookie ("BpCd")) Then
			frm1.txtBpCd.Value	=  ReadCookie ("BpCd")
			WriteCookie "BpCd",""
		Else
			'frm1.txtBpCd.Value	=  arrVal(5)
			frm1.txtBpCd.Value	=  arrVal(6)
		End If

		'frm1.txtBpNm.value		=  arrVal(6)
		frm1.txtBpNm.value		=  arrVal(7)

		If arrVal(6) = "" or arrVal(7) = Null Then
			frm1.txtDvFrDt.Text	=  ReadCookie ("DvFrDt")
			WriteCookie "DvFrDt",""
		Else
			'frm1.txtDvFrDt.Text		=  arrVal(7)
			frm1.txtDvFrDt.Text		=  arrVal(8)
		End If

		If arrVal(6) = "" or arrVal(7) = Null Then
			frm1.txtDvToDt.Text	=  ReadCookie ("DvToDt")
			WriteCookie "DvToDt",""
		Else
			'frm1.txtDvToDt.Text		=  arrVal(7)
			frm1.txtDvToDt.Text		=  arrVal(8)
		End If

		If Len(ReadCookie ("PurGrpCd")) Then
			frm1.txtPurGrpCd.Value	=  ReadCookie ("PurGrpCd")
			WriteCookie "PurGrpCd",""
		Else
			'frm1.txtPurGrpCd.Value	=  arrVal(8)
			frm1.txtPurGrpCd.Value	=  arrVal(9)
		End If

		'frm1.txtPurGrpNm.value 	=  arrVal(9)
		frm1.txtPurGrpNm.value 	=  arrVal(10)

		If Len(ReadCookie ("PoType")) Then
			frm1.txtPoType.Value	=  ReadCookie ("PoType")
			WriteCookie "PoType",""
		Else
			'frm1.txtPoType.Value	=  arrVal(10)
			frm1.txtPoType.Value	=  arrVal(11)
		End If

		'frm1.txtPoTypeNm.value	=  arrVal(11)
		frm1.txtPoTypeNm.value	=  arrVal(12)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function
'================================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
	  gMouseClickStatus = "SPCR"
   End If
End Sub
'================================================================================================================================
Function FncSplitColumn()

	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit

	iColumnLimit  = frm1.vspdData.MaxCols

	ACol = Frm1.vspdData.ActiveCol
	ARow = Frm1.vspdData.ActiveRow

	If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
	   Exit Function
	End If

	Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE

	ggoSpread.Source = Frm1.vspdData

	ggoSpread.SSSetSplit(ACol)

	Call SetActiveCell(frm1.vspdData,ACol,ARow,"M","X","X")

	Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH

End Function
'================================================================================================================================
Sub txtDvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDvFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDvFrDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtDvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDvToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDvToDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtDvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'================================================================================================================================
Sub txtDvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'================================================================================================================================
Sub vspdData_GotFocus()
	ggoSpread.Source = frm1.vspdData
End Sub
'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then
		 Exit Sub
	End If
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Set gActiveSpdSheet = frm1.vspdData

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"

	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If

		Exit Sub
	End If
	Call SetSpreadColumnValue("A",Frm1.vspdData, Col, Row)
End Sub
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)

	If OldLeft <> NewLeft Then
		Exit Sub
	End If


	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then		'☜: 재쿼리 체크 
		If lgPageNo <> "" Then															'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'================================================================================================================================
Function FncQuery()

	FncQuery = False
	Err.Clear

	with frm1
		if (UniConvDateToYYYYMMDD(.txtDvFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtDvToDt.text,Parent.gDateFormat,"")) And Trim(.txtDvFrDt.text) <> "" And Trim(.txtDvToDt.text) <> "" then
			Call DisplayMsgBox("17a003","X","납기일","X")
			Exit Function
		End if
	End with

	Call InitVariables

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	If DbQuery = False Then Exit Function

	FncQuery = True

	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncPrint()
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncExcel()
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncFind()
	Call parent.FncFind(parent.C_MULTI , False)
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncExit()
	FncExit = True
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function DbQuery()
	Dim strVal

	DbQuery = False

	Err.Clear
	If CheckRunningBizProcess = True Then
	   Exit Function
	End If

	Call LayerShowHide(1)

	With frm1
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.hdnPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
		strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.value)
		strVal = strVal & "&txtDvFrDt=" & Trim(.hdnDvFrDt.value)
		strVal = strVal & "&txtDvToDt=" & Trim(.hdnDvToDt.value)
		strVal = strVal & "&txtPurGrpCd=" & Trim(.hdnPurGrpCd.value)
		strVal = strVal & "&txtPoType=" & Trim(.hdnPoType.value)
		strVal = strVal & "&txtTrackNo=" & .hdnTrackNo.Value

	else
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
		strVal = strVal & "&txtDvFrDt=" & Trim(.txtDvFrDt.Text)
		strVal = strVal & "&txtDvToDt=" & Trim(.txtDvToDt.Text)
		strVal = strVal & "&txtPurGrpCd=" & Trim(.txtPurGrpCd.value)
		strVal = strVal & "&txtPoType=" & Trim(.txtPoType.value)
		strVal = strVal & "&txtTrackNo=" & Trim(.txtTrackNo.Value)

	end if

		strVal = strVal & "&lgPageNo="   & lgPageNo
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="	 & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

		Call RunMyBizASP(MyBizASP, strVal)

	End With

	DbQuery = True
	Call SetToolbar("1100000000011111")

End Function
'================================================================================================================================
Function DbQueryOk()

	lgBlnFlgChgValue = False
	lgSaveRow		= 1
	lgIntFlgMode = Parent.OPMD_UMODE

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement

End Function
'================================================================================================================================
