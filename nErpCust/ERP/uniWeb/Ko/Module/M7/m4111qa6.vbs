
Option Explicit

Dim lgIsOpenPop
Dim lgSaveRow
Dim IscookieSplit

Const BIZ_PGM_ID 		= "m4111qb6.asp"
Const BIZ_PGM_JUMP_ID 	= "m4111ma1"
Const C_MaxKey          = 31

'================================================================================================================================
 Sub InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = Parent.OPMD_CMODE
End Sub
'================================================================================================================================
Sub SetDefaultVal()
 	frm1.txtMvFrDt.Text	= StartDate
 	frm1.txtMvToDt.Text	= EndDate
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
 	Call SetZAdoSpreadSheet("M4111QA6","S","A","V20050529", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A")
End Sub
'================================================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub
''================================================================================================================================
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"
	arrParam(1) = "B_Plant"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	arrParam(4) = ""
	arrParam(5) = "����"

    arrField(0) = "Plant_Cd"
    arrField(1) = "Plant_NM"

    arrHeader(0) = "����"
    arrHeader(1) = "�����"

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
		Call DisplayMsgBox("17A002","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

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

	arrParam(0) = "�ŷ�ó"
	arrParam(1) = "B_Biz_Partner"
	arrParam(2) = Trim(frm1.txtBpCd.Value)
'	arrParam(3) = Trim(frm1.txtBpNm.Value)
	arrParam(4) = "BP_TYPE in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "
	arrParam(5) = "�ŷ�ó"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "�ŷ�ó"
    arrHeader(1) = "�ŷ�ó��"

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
Function OpenSlCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "â��"
	arrParam(1) = "B_STORAGE_LOCATION"
	arrParam(2) = Trim(frm1.txtSlCd.Value)
'	arrParam(3) = Trim(frm1.txtSlNm.Value)
	arrParam(4) = ""
	arrParam(5) = "â��"

    arrField(0) = "SL_CD"
    arrField(1) = "SL_NM"

    arrHeader(0) = "â��"
    arrHeader(1) = "â���"

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
Function OpenIoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�԰�����"
	arrParam(1) = "M_MVMT_TYPE"
	arrParam(2) = Trim(frm1.txtIoType.Value)
'	arrParam(3) = Trim(frm1.txtIoTypeNm.Value)
	arrParam(4) = "RCPT_FLG <> " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "�԰�����"

    arrField(0) = "IO_TYPE_CD"
    arrField(1) = "IO_TYPE_NM"

    arrHeader(0) = "�԰�����"
    arrHeader(1) = "�԰�������"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtIoTYpe.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtIoTYpe.Value = arrRet(0)
		frm1.txtIoTypeNm.Value = arrRet(1)
		frm1.txtIoTYpe.focus
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
'================================================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877

	If Kubun = 1 Then

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie CookieSplit , IscookieSplit
		if frm1.vspdData.ActiveRow > 0 then
			strTemp = ReadCookie(CookieSplit)
			If strTemp = "" then Exit Function
			arrVal = Split(strTemp, parent.gRowSep)
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			WriteCookie "MvmtNo" , arrVal(0)
			WriteCookie CookieSplit , ""
		end if
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then
	  'arrVal(2)�� ǰ��԰� 
	  If ReadCookie("From")="PO" Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)
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
			frm1.txtBpCd.Value	=  arrVal(5)
		End If

		frm1.txtBpNm.value		=  arrVal(6)

		WriteCookie "From",""

	  Else

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function

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
			frm1.txtBpCd.Value	=  arrVal(5)
		End If

		frm1.txtBpNm.value		=  arrVal(6)

		If arrVal(6) = "" or arrVal(7) = Null Then
			frm1.txtMvFrDt.Text	=  ReadCookie ("MvFrDt")
			WriteCookie "MvFrDt",""
		Else
			frm1.txtMvFrDt.Text		=  arrVal(7)
		End If

		If arrVal(6) = "" or arrVal(7) = Null Then
			frm1.txtMvToDt.Text	=  ReadCookie ("MvToDt")
			WriteCookie "MvToDt",""
		Else
			frm1.txtMvToDt.Text		=  arrVal(7)
		End If

		If Len(ReadCookie ("SlCd")) Then
			frm1.txtSlCd.Value	=  ReadCookie ("SlCd")
			WriteCookie "SlCd",""
		Else
			frm1.txtSlCd.Value	=  arrVal(9)
		End If

		frm1.txtSlNm.value 	=  arrVal(10)

		If Len(ReadCookie ("IoType")) Then
			frm1.txtIoType.Value	=  ReadCookie ("IoType")
			WriteCookie "IoType",""
		Else
			frm1.txtIoType.Value	=  arrVal(11)
		End If

		frm1.txtIoTypeNm.value	=  arrVal(12)
		
		frm1.txtTrackNo.value	=  arrVal(8)	

	  End if

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
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub
'================================================================================================================================
Sub txtMvFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtMvFrDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtMvToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtMvToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtMvToDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtMvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'================================================================================================================================
Sub txtMvToDt_KeyDown(KeyCode, Shift)
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

	IscookieSplit=""
	IscookieSplit = IscookieSplit & GetSpreadText(frm1.vspdData,GetKeyPos("A",1),Row,"X","X") & parent.gRowSep

End Sub
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)

	If OldLeft <> NewLeft Then
	    Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ 
		If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
		if (UniConvDateToYYYYMMDD(.txtMvFrDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtMvToDt.text,Parent.gDateFormat,"")) And Trim(.txtMvFrDt.text) <> "" And Trim(.txtMvToDt.text) <> "" then
			Call DisplayMsgBox("17a003","X","�԰���","X")
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

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim IntRetCD
	Dim iCalledAspName

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
    arrParam(2) = Trim(frm1.txtPlantCd.value)	'���� 
    arrParam(3) = ""	'��ǰ�� 
    arrParam(4) = ""	'���ֹ�ȣ 
    arrParam(5) = ""	'�߰� Where�� 

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
Function DbQuery()
	Dim strVal

    DbQuery = False

    Err.Clear
    If CheckRunningBizProcess = True Then
       Exit Function
    End If

    Call LayerShowHide(1)


    With frm1
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.hdnPlantCd.value)
    	strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.Value)
    	strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.Value)
    	strVal = strVal & "&txtMvFrDt=" & Trim(.hdnMvFrDt.value)
    	strVal = strVal & "&txtMvToDt=" & Trim(.hdnMvToDt.value)
    	strVal = strVal & "&txtSlCd=" & Trim(.hdnSlCd.value)
    	strVal = strVal & "&txtIoType=" & Trim(.hdnIoType.value)
		strVal = strVal & "&txtTrackNo=" & .hdnTrackNo.Value

    else
	    strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
    	strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
    	strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.Value)
    	strVal = strVal & "&txtMvFrDt=" & Trim(.txtMvFrDt.Text)
    	strVal = strVal & "&txtMvToDt=" & Trim(.txtMvToDt.Text)
    	strVal = strVal & "&txtSlCd=" & Trim(.txtSlCd.value)
    	strVal = strVal & "&txtIoType=" & Trim(.txtIoType.value)
		strVal = strVal & "&txtTrackNo=" & Trim(.txtTrackNo.Value)
    end if
        strVal = strVal & "&lgPageNo="   & lgPageNo
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)

    End With

    DbQuery = True
    Call SetToolbar("1100000000011111")

End Function
'================================================================================================================================
Function DbQueryOk()

    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    lgIntFlgMode = Parent.OPMD_UMODE

    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtPlantCd.focus
	End If

	Set gActiveElement = document.activeElement
End Function
