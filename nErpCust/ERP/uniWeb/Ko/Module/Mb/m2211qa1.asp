<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name		  : Procurement
'*  2. Function Name		:
'*  3. Program ID		   :
'*  4. Program Name		 :
'*  5. Program Desc		 :
'*  6. Comproxy List		:
'*  7. Modified date(First) :
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)	 : Byun Jee Hyun
'* 10. Modifier (Last)	  : Park Jin Uk
'* 11. Comment			  :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*							this mark(⊙) Means that "may  change"
'*							this mark(☆) Means that "must change"
'* 13. History			  :
'*							2000/12/09
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop
Dim lgSaveRow
Dim IsCookieSplit

Const BIZ_PGM_ID		= "M2211QB1.asp"
Const BIZ_PGM_JUMP_ID   = "M2211QA2"
Const C_MaxKey		  = 22

'================================================================================================================================
Dim StartDate, EndDate
StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate   = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'================================================================================================================================
Sub InitVariables()
	lgBlnFlgChgValue = False							   'Indicates that no value changed
	lgPageNo	 = ""								  'initializes Previous Key
	lgSortKey		= 1
	lgIntFlgMode = Parent.OPMD_CMODE
End Sub
'================================================================================================================================
Sub SetDefaultVal()
	With frm1
 		.txtDlvyFrDt.Text	= StartDate
 		.txtDlvyToDt.Text	= EndDate
		.txtPlantCd.value= parent.gPlant
		.txtPlantCd.focus
	End With
	Set gActiveElement = document.activeElement
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M2211QA1","G","A","V20040521", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
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
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenGroupPopup("A")
End Sub
'================================================================================================================================
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
		Call InitVariables
		Call InitSpreadSheet()
	End If
End Function
'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_Plant"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "Plant_CD"
	arrField(1) = "Plant_NM"

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
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(parent.UCN_PROTECTED) then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	lgIsOpenPop = True

	arrParam(0) = "모품목"
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"

	arrParam(2) = Trim(frm1.txtitemCd.Value)

	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(UCase(frm1.txtPlantCd.Value), "''", "S") & " "
	End if

	arrParam(5) = "모품목"

	arrField(0) = "B_Item.Item_Cd"
	arrField(1) = "B_Item.Item_NM"
	arrField(2) = "B_Plant.Plant_Cd"
	arrField(3) = "B_Plant.Plant_NM"

	arrHeader(2) = "공장"
	arrHeader(3) = "공장명"

	arrHeader(0) = "모품목"
	arrHeader(1) = "모품목명"

	iCalledAspName = AskPRAspName("M1111PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M1111PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField, arrHeader), _
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
Function OpenSl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	if UCase(frm1.txtSlCd.className) = UCase(parent.UCN_PROTECTED) then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "출고창고"
	arrParam(1) = "B_Storage_location"
	arrParam(2) = Trim(frm1.txtSlCd.Value)
	arrParam(4) = ""
	arrParam(5) = "출고창고"

	arrField(0) = "B_Storage_location.Sl_Cd"
	arrField(1) = "B_Storage_location.Sl_Nm"

	arrHeader(0) = "출고창고"
	arrHeader(1) = "출고창고명"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSlCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSlCd.Value = arrRet(0)
		frm1.txtSlNm.Value = arrRet(1)
		frm1.txtSlCd.focus
		Set gActiveElement = document.activeElement
	End If

End Function
'================================================================================================================================
Function OpenPoNo()

	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrParam(5)

	If lgIsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	iCalledAspName = AskPRAspName("M3111PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoNo.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoNo.Value = arrRet(0)
		frm1.txtPoNo.focus
		Set gActiveElement = document.activeElement
	End If

End Function
'================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True  Or UCase(frm1.txtSpplCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공급처"
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = Trim(frm1.txtSpplCd.Value)
	'arrParam(3) = Trim(frm1.txtSpplNm.Value)
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "공급처"

	arrField(0) = "BP_Cd"
	arrField(1) = "BP_NM"

	arrHeader(0) = "공급처"
	arrHeader(1) = "공급처명"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSpplCd.Value = arrRet(0)
		frm1.txtSpplNm.Value = arrRet(1)
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function
'================================================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = "Cookiekeym2211qa1"

	If Kubun = 1 Then

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie CookieSplit , IsCookieSplit

		If Len(Trim(frm1.txtPlantCd.value)) Then
			WriteCookie "PlantCd",Trim(frm1.txtPlantCd.value)
		Else
			WriteCookie "PlantCd",""
		End If

		If Len(Trim(frm1.txtItemCd.value)) Then
			WriteCookie "ItemCd",Trim(frm1.txtItemCd.value)
		Else
			WriteCookie "ItemCd",""
		End If

		If Len(Trim(frm1.txtSlCd.value)) Then
			WriteCookie "SlCd",Trim(frm1.txtSlCd.value)
		Else
			WriteCookie "SlCd",""
		End If

		If Len(Trim(frm1.txtDlvyFrDt.text)) Then
			WriteCookie "DlvyFrDt",Trim(frm1.txtDlvyFrDt.text)
		Else
			WriteCookie "DlvyFrDt",""
		End If

		If Len(Trim(frm1.txtDlvyToDt.text)) Then
			WriteCookie "DlvyToDt",Trim(frm1.txtDlvyToDt.text)
		Else
			WriteCookie "DlvyToDt",""
		End If

		If Len(Trim(frm1.txtSpplCd.value)) Then
			WriteCookie "SpplCd",Trim(frm1.txtSpplCd.value)
		Else
			WriteCookie "SpplCd",""
		End If

		If  frm1.rdoUseflg(0).checked=true then
		  WriteCookie "rdoUseflg","A"
		elseif frm1.rdoUseflg(1).checked=true then
		  WriteCookie "rdoUseflg","F"
		 else
		 WriteCookie "rdoUseflg","C"
		 end If

		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, gRowSep)

		'If arrVal(0) = "" Then Exit Function

		Dim iniSep

		frm1.txtPlantCd.value =  arrVal(0)
		frm1.txtPlantNm.value =  arrVal(1)
		frm1.txtItemCd.Value =  arrVal(2)
		frm1.txtItemCd.Value =  arrVal(3)
		frm1.txtSlCd.value =  arrVal(4)
		frm1.txtSlNm.value = arrVal(5)
		frm1.txtDlvyFrDt.Text = arrVal(6)
		frm1.txtDlvyToDt.text =  arrVal(7)
		frm1.txtSpplCd.value = arrVal(8)
		frm1.txtSpplNm.value = arrVal(9)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

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
Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call SetToolbar("1100000000001111")
	Call CookiePage(0)

	Set gActiveElement = document.activeElement
End Sub
'================================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
	  gMouseClickStatus = "SPCR"
   End If
End Sub
'================================================================================================================================
Sub FncSplitColumn()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Dim iRowSep
	Dim ii

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

	Call SetSpreadColumnValue("A",frm1.vspdData, Col, Row)

	IsCookieSplit = ""
	iRowSep = Parent.gRowSep

	For ii = 1 to C_MaxKey
		IsCookieSplit = IsCookieSplit & Trim(GetSpreadText(frm1.vspdData,GetKeyPos("A",ii),Row,"X","X")) & iRowSep
	Next

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
Sub txtDlvyFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDlvyFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDlvyFrDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtDlvyToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDlvyToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDlvyToDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtDlvyFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'================================================================================================================================
Sub txtDlvyToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'================================================================================================================================
Function FncQuery()

	FncQuery = False

	Err.Clear

	With frm1
		if (UniConvDateToYYYYMMDD(.txtDlvyFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtDlvyToDt.text,Parent.gDateFormat,"")) And Trim(.txtDlvyFrDt.text) <> "" And Trim(.txtDlvyToDt.text) <> "" then
			Call DisplayMsgBox("17a003", "X","필요일", "X")
			Exit Function
		End if
	End with

	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	Call InitVariables

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

	If LayerShowHide(1) = False Then Exit Function

	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtPlantCd=" & Trim(.hdnPlantCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
			strVal = strVal & "&txtSlCd=" & Trim(.hdnSlCd.value)
			strVal = strVal & "&txtDlvyFrDt=" & Trim(.hdnDlvyFrDt.value)
			strVal = strVal & "&txtDlvyToDt=" & Trim(.hdnDlvyToDt.value)
			strVal = strVal & "&txtSpplCd=" & Trim(.hdnSpplCd.Value)
			strVal = strVal & "&rdoUseflg=" & Trim(.hdnrdoUseflg.value)
			strVal = strVal & "&txtTrackNo=" & .hdnTrackNo.Value
			strVal = strVal & "&rdoClsflg=" & Trim(.hdnrdoClsflg.value)
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtSlCd=" & Trim(.txtSlCd.value)
			strVal = strVal & "&txtDlvyFrDt=" & Trim(.txtDlvyFrDt.Text)
			strVal = strVal & "&txtDlvyToDt=" & Trim(.txtDlvyToDt.Text)
			strVal = strVal & "&txtSpplCd=" & Trim(.txtSpplCd.Value)
			strVal = strVal & "&txtTrackNo=" & Trim(.txtTrackNo.Value)

			if .rdoUseflg(0).checked=true then
				strVal = strVal & "&rdoUseflg=" & Trim(.rdoUseflg(0).value)
			elseif .rdoUseflg(1).checked=true then
				strVal = strVal & "&rdoUseflg=" & Trim(.rdoUseflg(1).value)
			else
				strVal = strVal & "&rdoUseflg=" & Trim(.rdoUseflg(2).value)
			end if
			
			if .rdoClsflg(0).checked=true then
				strVal = strVal & "&rdoClsflg=" & Trim(.rdoClsflg(0).value)
			elseif .rdoClsflg(1).checked=true then
				strVal = strVal & "&rdoClsflg=" & Trim(.rdoClsflg(1).value)
			else
				strVal = strVal & "&rdoClsflg=" & Trim(.rdoClsflg(2).value)
			end if

		End if

		strVal = strVal & "&lgPageNo="   & lgPageNo
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="	 & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		Call RunMyBizASP(MyBizASP, strVal)

	End With

	DbQuery = True
	Call SetToolbar("1100000000011111")

End Function
'================================================================================================================================
Function DbQueryOk()
	lgIntFlgMode = Parent.OPMD_UMODE

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>예약자재집계</font></td>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd"  SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
														   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>모품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="모품목" NAME="txtItemCd"  SIZE=10 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
														   <INPUT TYPE=TEXT ALT="모품목" NAME="txtItemNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>출고창고</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="출고창고" NAME="txtSlCd"  SIZE=10 MAXLENGTH=7 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSl()">
														   <INPUT TYPE=TEXT ALT="출고창고" NAME="txtSlNm" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>필요일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=필요일 NAME="txtDlvyFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
												   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=필요일 NAME="txtDlvyToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											</tr>
										</table>
									</TD>
								<TR>
									<TD CLASS="TD5" NOWRAP>지급구분</TD>
									 <TD CLASS="TD6" NOWRAP><INPUT TYPE=radio ALT="지급구분" class="radio" NAME="rdoUseflg" checked value = "A" tag="1X">
														   <label for="rdoUseflg">전체</label>
														   <INPUT TYPE=radio ALT="지급구분" class="radio" NAME="rdoUseflg" value = "F" tag="1X">
														   <label for="rdoUseflg">무상</label>
														   <INPUT TYPE=radio ALT="지급구분" class="radio" NAME="rdoUseflg" value = "C" tag="1X">
														   <label for="rdoUseflg">유상</label></TD>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="공급처" NAME="txtSpplCd"  MAXLENGTH=10 SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
														   <INPUT TYPE=TEXT AlT="공급처" ID="txtSpplNm" NAME="arrCond" tag="14X"></TD>
								</TR>
								<tr>
									<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="Tracking No." NAME="txtTrackNo" SIZE=34 MAXLENGTH=25  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingNo()"></TD>
									<TD CLASS="TD5" NOWRAP>마감여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio ALT="마감여부" class="radio" NAME="rdoClsflg" checked value = "A" tag="1X">
														   <label for="rdoClsflg">전체</label>
														   <INPUT TYPE=radio ALT="마감여부" class="radio" NAME="rdoClsflg" value = "N" tag="1X">
														   <label for="rdoClsflg">미마감</label>
														   <INPUT TYPE=radio ALT="마감여부" class="radio" NAME="rdoClsflg" value = "Y" tag="1X">
														   <label for="rdoClsflg">마감</label></TD>
								</tr>

							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</tr>
	<TR HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">예약자재상세현황</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="HconItem_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HValid_from_dt" tag="24">
<INPUT TYPE=HIDDEN NAME="HconCurrency" tag="24">
<INPUT TYPE=HIDDEN NAME="HconDeal_type" tag="24">
<INPUT TYPE=HIDDEN NAME="HconPay_terms" tag="24">
<INPUT TYPE=HIDDEN NAME="HconSales_unit" tag="24">

<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSlCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDlvyFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDlvyToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSpplCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnrdoUseflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnrdoClsflg" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
