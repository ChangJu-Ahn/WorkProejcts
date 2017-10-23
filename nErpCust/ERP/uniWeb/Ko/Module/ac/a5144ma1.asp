<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5144MA1
'*  4. Program Name         : 입력경로별계정조회 
'*  5. Program Desc         : Query of Account Code
'*  6. Component List       : ADO
'*  7. Modified date(First) : 2003/06/05
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : Jung Sung Ki
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================================================================
Dim lgIsOpenPop
Dim IsOpenPop                                               '☜: Popup status
Dim lgMark                                                  '☜: 마크 
Dim  gSelframeFlg
Dim lgPageNo2
Dim lgIntFlgMode2
'==========================================================================================
Const BIZ_PGM_ID		= "A5144MB1.asp"
Const BIZ_PGM_ID2		= "A5144MB2.asp"
Const BIZ_PGM_ID3		= "A5144MB3.asp"

Const C_MaxKey          = 5

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

'==========================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False
    lgPageNo		= ""
    lgPageNo2		= ""
    lgSortKey		= 1
    lgIntFlgMode     = Parent.OPMD_CMODE
    lgIntFlgMode2     = Parent.OPMD_CMODE
End Sub

'==========================================================================================
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

	EndDate = "<%=GetSvrDate%>"
	'Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	StartDate	= UNIGetFirstDay(UNIDateAdd("m", -1, EndDate, parent.gServerDateFormat),Parent.gServerDateFormat)
	EndDate		= UNIGetLastDay(StartDate , Parent.gServerDateFormat)

	frm1.txtFromGlDt.Text   = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtToGlDt.Text     = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtIssuedDt1.Text     = ""
	frm1.txtIssuedDt2.Text     = ""

	frm1.txtFromGlDt2.Text   = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtToGlDt2.Text     = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
	frm1.txtIssuedDt21.Text     = ""
	frm1.txtIssuedDt22.Text     = ""

	frm1.txtFromGlDt.focus

End Sub

'==========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'==========================================================================================
Sub InitSpreadSheet()

		Call SetZAdoSpreadSheet("A5144MA1","S","A","V20021211",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
		Call SetSpreadLock("A")
		Call SetZAdoSpreadSheet("A5144MA102","S","B","V20021212",parent.C_SORT_DBAGENT,frm1.vspdData3, C_MaxKey, "X","X")
		Call SetSpreadLock("B")
		Call SetZAdoSpreadSheet("A5144MA101","S","C","V20030620",parent.C_SORT_DBAGENT,frm1.vspdData1, C_MaxKey, "X","X")
		Call SetSpreadLock("C")
		Call SetZAdoSpreadSheet("A5144MA101","S","D","V20030620",parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
		Call SetSpreadLock("D")
End Sub



'==========================================================================================
Sub SetSpreadLock(ByVal pOpt)
    Select Case UCase(Trim(pOpt))
		Case "A"
			With frm1.vspdData
				.ReDraw = False 
					ggoSpread.SpreadLockWithOddEvenRowColor()
				.ReDraw = True
			End With
		Case "B"
			With frm1.vspdData3
				.ReDraw = False 
					ggoSpread.SpreadLockWithOddEvenRowColor()
				.ReDraw = True
			End With
		Case "C"
			With frm1.vspdData1
				.ReDraw = False 
					ggoSpread.SpreadLockWithOddEvenRowColor()
				.ReDraw = True
			End With
		Case "D"
			With frm1.vspdData2
				.ReDraw = False 
					ggoSpread.SpreadLockWithOddEvenRowColor()
				.ReDraw = True
			End With
	End Select

End Sub



'==========================================================================================
Sub InitComboBox()	
	Err.clear
End Sub


'==========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	'frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 0,1
			arrParam(0) = "사업장 팝업"						' 팝업 명칭 
			arrParam(1) = "B_Biz_AREA"							' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "사업장코드"

			arrField(0) = "BIZ_AREA_CD"								' Field명(0)
			arrField(1) = "BIZ_AREA_NM"								' Field명(1)

			arrHeader(0) = "사업장코드"							' Header명(0)
			arrHeader(1) = "사업장명"							' Header명(1)
		Case 2, 3
			arrParam(0) = "전표입력경로팝업"
			arrParam(1) = "B_MINOR A, B_MINOR B , B_CONFIGURATION C"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = " A.MINOR_CD = C.MINOR_CD AND A.MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  AND B.MAJOR_CD=" & FilterVar("B0001", "''", "S") & "  AND C.MAJOR_CD=" & FilterVar("A1001", "''", "S") & "  AND B.MINOR_CD=C.REFERENCE "
			arrParam(5) = "전표입력경로코드"

			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			arrField(2) = "C.REFERENCE"
			arrField(3) = "B.MINOR_NM"

			arrHeader(0) = "전표입력경로코드"
			arrHeader(1) = "전표입력경로명"
			arrHeader(2) = "모듈코드"
			arrHeader(3) = "모듈명"


		Case 5
			arrParam(0) = "모듈코드팝업"
			arrParam(1) = "B_MINOR"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = " MINOR_CD >= " & FilterVar("A", "''", "S") & "  AND MAJOR_CD = " & FilterVar("B0001", "''", "S") & " "
			arrParam(5) = "모듈코드"

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"

			arrHeader(0) = "모듈코드"
			arrHeader(1) = "모듈명"
		Case Else
		Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
		Case 0					'사업장1
			frm1.txtBizAreaCd.focus
		Case 1					'사업장2
			frm1.txtBizAreaCd2.focus
		Case 2					'입력경로1
			frm1.txtInputType.focus
		Case 3					'입력경로2
			frm1.txtInputType2.focus
		Case 5					'모듈코드 
			frm1.txtMocd.focus
		End Select
		Exit Function
	Else
		Select Case iWhere
		Case 0					'사업장1
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = Trim(arrRet(0))
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 1					'사업장2
			frm1.txtBizAreaCd2.focus
			frm1.txtBizAreaCd2.value = Trim(arrRet(0))
			frm1.txtBizAreaNm2.value = arrRet(1)
		Case 2					'입력경로 
			frm1.txtInputType.focus
			frm1.txtInputType.value = Trim(arrRet(0))
			frm1.txtInputTypeNM.value = arrRet(1)
		Case 3					'입력경로 
			frm1.txtInputType2.focus
			frm1.txtInputType2.value = Trim(arrRet(0))
			frm1.txtInputTypeNM2.value = arrRet(1)
		Case 5					'입력경로 
			frm1.txtMocd.focus
			frm1.txtMocd.value = Trim(arrRet(0))
			frm1.txtMoNm.value = arrRet(1)
		End Select
	End If

End Function

'==========================================================================================
'	Name : OpenBpCd()
'	Description : Bp Cd PopUp
'==========================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처 팝업"
	arrParam(1) = "B_BIZ_PARTNER"
	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래처코드"

    arrField(0) = "BP_CD"
    arrField(1) = "BP_NM"

    arrHeader(0) = "거래처코드"
    arrHeader(1) = "거래처명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtBpCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBpCd(arrRet)
	End If	
End Function

'==========================================================================================
'	Name : SetBpCd()
'	Description : Bp Cd Popup에서 Return되는 값 setting
'==========================================================================================
Function SetBpCd(byval arrRet)
	frm1.txtBpCd.Value    = arrRet(0)
	frm1.txtBpNm.Value    = arrRet(1)
	lgBlnFlgChgValue = True
End Function


'==========================================================================================
'	Name : OpenAcctCd()
'	Description : Account PopUp
'==========================================================================================
Function OpenAcctCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"									' 팝업 명칭 
	arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
	arrParam(2) = strCode											' Code Condition
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"					' Where Condition
'	If frm1.hAcctbalfg.Value <> "" and iWhere = 3 Then
'		arrParam(4) = arrParam(4) & " AND A_Acct.bal_fg = " & Filtervar(frm1.hAcctbalfg.Value, "''", "S")
'	End If
	arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

	arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
	arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
	arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
'	arrField(4) = "HH" & parent.gColSep & "A_Acct.bal_fg"									' Field명(3)

	arrHeader(0) = "계정코드"									' Header명(0)
	arrHeader(1) = "계정코드명"									' Header명(1)
	arrHeader(2) = "그룹코드"									' Header명(2)
	arrHeader(3) = "그룹명"										' Header명(3)
'	arrHeader(4) = "차대구분"										' Header명(3)


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	Select case iWhere
	case 1
		frm1.txtBizAreaCd.focus
	case 2
		frm1.txtAcctCd1.focus
	case 3
		frm1.txtAcctCd2.focus
	End select
	If arrRet(0) = "" Then

		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If

End Function
'==========================================================================================
Function OpenPopupAcct()

	Dim arrRet
	Dim arrParam(6)
	Dim IntRetCD
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5144ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5144ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	If gSelframeFlg = TAB1 Then

		arrParam(0) = Trim(GetKeyPosVal("A", 1))	'계정코드 
		arrParam(1) = Trim(GetKeyPosVal("A", 2))	'입력경로 
		arrParam(2) = frm1.txtFromGlDt.Text			'From Date
		arrParam(3) = frm1.txtToGlDt.Text			'To Date
		arrParam(4) = Trim(frm1.txtBizAreaCd.value)	'사업장 
		
	ElseIf gSelframeFlg = TAB3 Then 
		arrParam(0) = Trim(GetKeyPosVal("B", 1))	'계정코드 
		arrParam(1) = Trim(GetKeyPosVal("B", 2))	'입력경로 
		arrParam(2) = frm1.txtFromGlDt2.Text		'From Date
		arrParam(3) = frm1.txtToGlDt2.Text			'To Date
		arrParam(4) = Trim(frm1.txtBizAreaCd2.value)'사업장 
	End If
		arrParam(5) = frm1.hOrgChangeId.value	'입력경로 
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	Dim gPos

	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       Case "vspdData3"
	            gPos = "B"
	       End Select

	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "2")
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()
   End If
End Function



'==========================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877

	If Kubun = 1 Then

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie "PoNo" , lsPoNo
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function

		Dim iniSep

'--------------- 개발자 coding part(실행로직,Start)---------------------------------------------------
		 '자동조회되는 조건값과 검색조건부 Name의 Match
		For iniSep = 0 To UBound(arrVal) -1
			Select Case UCase(Trim(arrVal(iniSep)))
			Case UCase("발주형태")
				frm1.txtPoType.value =  arrVal(iniSep + 1)
			Case UCase("발주형태명")
				frm1.txtPoTypeNm.value =  arrVal(iniSep + 1)
			Case UCase("공급처")
				frm1.txtSpplCd.value =  arrVal(iniSep + 1)
			Case UCase("공급처명")
				frm1.txtSpplNm.value =  arrVal(iniSep + 1)
			Case UCase("구매그룹")
				frm1.txtPurGrpCd.value =  arrVal(iniSep + 1)
			Case UCase("구매그룹명")
				frm1.txtPurGrpNm.value =  arrVal(iniSep + 1)
			Case UCase("품목")
				frm1.txtItemCd.value =  arrVal(iniSep + 1)
			Case UCase("품목명")
				frm1.txtItemNm.value =  arrVal(iniSep + 1)
			Case UCase("Tracking No.")
				frm1.txtTrackNo.value =  arrVal(iniSep + 1)
			End Select
		Next
'--------------- 개발자 coding part(실행로직,End)---------------------------------------------------

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'==========================================================================================
Sub Form_Load()
    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call InitComboBox()

	Call FncSetToolBar("New")
    frm1.txtNDrAmt.allownull = False
    frm1.txtNCrAmt.allownull = False
    frm1.txtNDrAmt2.allownull = False
    frm1.txtNCrAmt2.allownull = False
	Call ClickTab1()

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'==========================================================================================
Sub txtAcctCd1_onChange()
'	If Trim(frm1.txtAcctCd1.value) <> "" Then
'		Call CommonQueryRs("BAL_FG", "A_ACCT", "ACCT_CD = " & Filtervar(Trim(frm1.txtAcctCd1.value), "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
'		frm1.hAcctbalfg.value = Replace(lgF0, chr(11), "")
'	Else
'		frm1.txtAcctNm.value = ""
'		frm1.hAcctbalfg.value = ""
'	End If
'	frm1.txtAcctCd2.value = ""
'	frm1.txtAcctNm2.value = ""
	
End Sub

'==========================================================================================
' Tab 1
'==========================================================================================

Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtFromGlDt.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtFromGlDt.Focus
    End If
End Sub

Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtToGlDt.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtToGlDt.Focus
    End If
End Sub

Sub txtIssuedDt1_DblClick(Button)
    If Button = 1 Then
       frm1.txtIssuedDt1.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtIssuedDt1.Focus
    End If
End Sub

Sub txtIssuedDt2_DblClick(Button)
    If Button = 1 Then
       frm1.txtIssuedDt2.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtIssuedDt2.Focus
    End If
End Sub


Sub txtFromGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtToGlDt.focus
	   Call FncQuery
	End If
End Sub

Sub txtToGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtFromGlDt.focus
	   Call FncQuery
	End If
End Sub

Sub txtIssuedDt1_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtIssuedDt2.focus
	   Call FncQuery
	End If
End Sub

Sub txtIssuedDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtIssuedDt1.focus
	   Call FncQuery
	End If
End Sub

'==========================================================================================
' Tab 2
'==========================================================================================
Sub txtFromGlDt2_DblClick(Button)
    If Button = 1 Then
       frm1.txtFromGlDt2.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtFromGlDt2.Focus
    End If
End Sub

Sub txtToGlDt2_DblClick(Button)
    If Button = 1 Then
       frm1.txtToGlDt2.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtToGlDt2.Focus
    End If
End Sub

Sub txtIssuedDt21_DblClick(Button)
    If Button = 1 Then
       frm1.txtIssuedDt21.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtIssuedDt21.Focus
    End If
End Sub

Sub txtIssuedDt22_DblClick(Button)
    If Button = 1 Then
       frm1.txtIssuedDt22.Action = 7
       Call SetFocusToDocument("M")
       frm1.txtIssuedDt22.Focus
    End If
End Sub

Sub txtFromGlDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtToGlDt2.focus
	   Call FncQuery
	End If
End Sub

Sub txtToGlDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtFromGlDt2.focus
	   Call FncQuery
	End If
End Sub

Sub txtIssuedDt21_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtIssuedDt22.focus
	   Call FncQuery
	End If
End Sub

Sub txtIssuedDt22_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		frm1.txtIssuedDt21.focus
	   Call FncQuery
	End If
End Sub

Sub txtBizAreaCd2_onKeyPress()
    If window.event.keycode = 13 Then
        Call fncQuery()
    End If
End Sub

Sub txtInputType2_onKeyPress()
    If window.event.keycode = 13 Then
        Call fncQuery()
    End If
End Sub

Sub txtMoCd_onKeyPress()
    If window.event.keycode = 13 Then
        Call fncQuery()
    End If
End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
	If Row < 1 Then Exit Sub

	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub


Sub vspdData_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then
    End If
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub



'==========================================================================================
Sub vspdData3_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SP2C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData3

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData3
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
	If Row < 1 Then Exit Sub

	Call SetSpreadColumnValue("B", frm1.vspdData3, Col, Row)
End Sub


'==========================================================================================
Sub vspdData3_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then
    End If
End Sub

'==========================================================================================
Sub vspdData3_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then
    	If lgPageNo2 <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub

'==========================================================================================
'   Event Name : txtAmtFr_Keypress
'   Event Desc : 
'==========================================================================================
Sub txtAmtFr_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub
'==========================================================================================
'   Event Name : txtAmtTo_Keypress
'   Event Desc : 
'==========================================================================================
Sub txtAmtTo_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub

'==========================================================================================
Function FncQuery() 

    FncQuery = False
    Err.Clear
    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
	Select CAse gSelframeFlg
	Case TAB1
		If Not chkField(Document, "1") Then
		   Exit Function
		End If
		If CompareDateByFormat(frm1.txtFromGlDt.Text, frm1.txtToGlDt.Text, frm1.txtFromGlDt.Alt, frm1.txtToGlDt.Alt, _
							"970025", frm1.txtFromGlDt.UserDefinedFormat, parent.gComDateType, true) = False Then
				frm1.txtFromGlDt.focus
				Exit Function
		End if
		If CompareDateByFormat(frm1.txtIssuedDt1.Text, frm1.txtIssuedDt2.Text, frm1.txtIssuedDt1.Alt, frm1.txtIssuedDt2.Alt, _
							"970025", frm1.txtIssuedDt1.UserDefinedFormat, parent.gComDateType, true) = False Then
				frm1.txtIssuedDt1.focus
				Exit Function
		End if
		Call ggoOper.ClearField(Document, "2")
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
	Case TAB3
		If Not chkField(Document, "3") Then
		   Exit Function
		End If
		If CompareDateByFormat(frm1.txtFromGlDt2.Text, frm1.txtToGlDt2.Text, frm1.txtFromGlDt2.Alt, frm1.txtToGlDt2.Alt, _
							"970025", frm1.txtFromGlDt2.UserDefinedFormat, parent.gComDateType, true) = False Then
				frm1.txtFromGlDt2.focus
				Exit Function
		End if
		If CompareDateByFormat(frm1.txtIssuedDt21.Text, frm1.txtIssuedDt22.Text, frm1.txtIssuedDt21.Alt, frm1.txtIssuedDt22.Alt, _
							"970025", frm1.txtIssuedDt21.UserDefinedFormat, parent.gComDateType, true) = False Then
				frm1.txtIssuedDt21.focus
				Exit Function
		End if
		Call ggoOper.ClearField(Document, "4")
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.ClearSpreadData
	Case TAB2
		If frm1.txtFromGlDt1.Text = "" Then Exit Function
	End Select

	Call FncSetToolBar("New")
    Call DbQuery

    FncQuery = True
End Function


'==========================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function


'==========================================================================================
Function FncExcel()
	Call parent.FncExport(parent.C_MULTI)
End Function


'==========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function

'==========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub


'==========================================================================================
Function FncExit()
    FncExit = True
End Function

'==========================================================================================
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'==========================================================================================
Function SetReturnVal(ByVal arrRet,ByVal field_fg) 
	With frm1	
		Select case field_fg
			case 1
				.txtBizAreaCd.Value		= Trim(arrRet(0))
				.txtBizAreaNm.Value		= Trim(arrRet(1))
			case 2
				.txtAcctCd1.Value		= Trim(arrRet(0))
				.txtAcctNm1.Value		= Trim(arrRet(1))
				.txtAcctCd2.Value		= Trim(arrRet(0))
				.txtAcctNm2.Value		= Trim(arrRet(1))
			case 3
				.txtAcctCd2.Value		= Trim(arrRet(0))
				.txtAcctNm2.Value		= Trim(arrRet(1))
		End select
	End With

End Function


'==========================================================================================
Function DbQuery()
	Dim strVal, strZeroFg

    DbQuery = False

    Err.Clear
	Call LayerShowHide(1)

    With frm1

	Select Case gSelframeFlg
	Case TAB1
		If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
			strVal = BIZ_PGM_ID & "?txtFromGlDt=" & UniConvDateToYYYYMMDD(frm1.txtFromGlDt.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtToGlDt=" & UniConvDateToYYYYMMDD(frm1.txtToGlDt.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtIssuedDt=" & UniConvDateToYYYYMMDD(frm1.txtIssuedDt1.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtIssuedDt2=" & UniConvDateToYYYYMMDD(frm1.txtIssuedDt2.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd1.Value)
			strVal = strVal & "&txtAcctCd2=" & Trim(.txtAcctCd2.Value)
			strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.Value)
		Else
			strVal = BIZ_PGM_ID & "?txtFromGlDt=" & .htxtFromGlDt.value
			strVal = strVal & "&txtToGlDt=" & .htxtToGlDt.value
			strVal = strVal & "&txtIssuedDt=" & .htxtIssuedDt1.value
			strVal = strVal & "&txtIssuedDt2=" & .htxtIssuedDt2.value
			strVal = strVal & "&txtAcctCd=" & .htxtAcctCd1.Value
			strVal = strVal & "&txtAcctCd2=" & Trim(.htxtAcctCd2.Value)
			strVal = strVal & "&txtBizAreaCd=" & Trim(.htxtBizAreaCd.Value)
        End If

		strVal = strVal & "&txtAcctCd_Alt=" & Trim(.txtAcctCd1.Alt)
		strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	
	Case TAB2

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

			strVal = BIZ_PGM_ID3 & "?txtFromGlDt=" & UniConvDate(frm1.txtFromGlDt1.Text)
			strVal = strVal & "&txtToGlDt=" & UniConvDate(frm1.txtToGlDt1.Text)
			strVal = strVal & "&txtIssuedDt=" & frm1.txtIssuedDtFR.Text
			strVal = strVal & "&txtIssuedDt2=" & frm1.txtIssuedDtTO.Text
			strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd1.Value)
			strVal = strVal & "&txtAcctCd=" & .txtAcctCd.value

		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("C")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("C")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("C"))

	Case TAB3

		If lgIntFlgMode2  <> Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID2 & "?txtFromGlDt=" & UniConvDateToYYYYMMDD(frm1.txtFromGlDt2.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtToGlDt=" & UniConvDateToYYYYMMDD(frm1.txtToGlDt2.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtIssuedDt=" & UniConvDateToYYYYMMDD(frm1.txtIssuedDt21.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtIssuedDt2=" & UniConvDateToYYYYMMDD(frm1.txtIssuedDt22.Text,parent.gDateFormat,"")
			strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd2.Value)
			strVal = strVal & "&txtInputType=" & .txtInputType2.value
			'strVal = strVal & "&txtMocd=" & .txtMocd.value
		Else
			strVal = BIZ_PGM_ID2 & "?txtFromGlDt=" & .htxtFromGlDt2.value
			strVal = strVal & "&txtToGlDt=" & .htxtToGlDt2.value
			strVal = strVal & "&txtIssuedDt=" & .htxtIssuedDt21.value
			strVal = strVal & "&txtIssuedDt2=" & .htxtIssuedDt22.value
			strVal = strVal & "&txtBizAreaCd=" & Trim(.htxtBizAreaCd2.Value)
			strVal = strVal & "&txtInputType=" & .htxtInputType.value
			'strVal = strVal & "&txtMocd=" & .htxtMocd.value
        End If

		strVal = strVal & "&lgPageNo2="   & lgPageNo2                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
	End Select

	Call RunMyBizASP(MyBizASP, strVal)

    End With

    DbQuery = True
End Function



'==========================================================================================
Function DbQueryOk()
	Call FncSetToolBar("Query")
	frm1.vspdData.focus
	CALL vspdData_Click(1, 1)
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function DbQueryOk2()
    lgIntFlgMode2     = Parent.OPMD_UMODE
	Call FncSetToolBar("Query")
	frm1.vspdData3.focus
	CALL vspdData3_Click(1, 1)
	Set gActiveElement = document.ActiveElement
End Function

Function DbQueryOk3()
	frm1.vspdData1.focus
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
'툴바버튼 세팅 
'==========================================================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function

'========================================================================================
' Function Name : ClickTab1
' Function Desc : This function tab1 click
'========================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB1
	frm1.txtFromGlDt.focus
	'Call SetDefaultVal()
	userview1.style.display = ""
	userview2.style.display = "NONE"

End Function

'========================================================================================
' Function Name : ClickTab2
' Function Desc : This function tab2 click
'========================================================================================
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function

	userview1.style.display = "NONE"
	userview2.style.display = ""
	userview2.disabled = True

	If gSelframeFlg = TAB1 and frm1.vspddata.ActiveRow > 0 then
		With frm1
			.txtFromGlDt1.text = .txtFromGlDt.Text
			.txtToGlDt1.Text = .txtToGlDt.Text
			.txtIssuedDtFR.Text = .txtIssuedDt1.Text
			.txtIssuedDtTO.Text = .txtIssuedDt2.Text
			.txtBizAreaCd1.Value = .txtBizAreaCd.Value
			.txtBizAreaNm1.Value = .txtBizAreaNm.Value
			ggoSpread.Source = .vspdData
			.vspdData.row = .vspddata.ActiveRow
			.vspddata.col = GetKeyPos("A", 1)
			.txtAcctCd.value = .vspdData.Text
			.vspddata.col = GetKeyPos("A", 3)
			.txtAcctNm.value = .vspdData.Text
		End With
		Call changeTabs(TAB2)	 '~~~ 두번째 Tab
		gSelframeFlg = TAB2
		Call DbQuery()
	Else
		Call changeTabs(TAB2)	 '~~~ 두번째 Tab
		gSelframeFlg = TAB2
		Exit Function
	End If
	
End Function

'========================================================================================
' Function Name : ClickTab3
' Function Desc : This function tab2 click
'========================================================================================
Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3)	 '~~~ 두번째 Tab
	gSelframeFlg = TAB3
	frm1.txtFromGlDt2.focus
	userview1.style.display = ""
	userview2.style.display = "NONE"

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정별입력경로조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>거래유형별조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>입력경로별계정조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*">&nbsp;</TD>
					<TD WIDTH="*" align=right ID=userview1><A HREF="VBSCRIPT:OpenPopupAcct()">계정별보조부조회</A>&nbsp;</td>
					<TD WIDTH="*" align=right ID=userview2>계정별보조부조회&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
		<DIV ID="TabDiv"  SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>회계일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5144ma1_fpDateTime1_txtFromGlDt.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/a5144ma1_fpDateTime2_txtToGlDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>사업장코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtBizAreaCd.value,0)"> <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=18 tag="24X" ALT="사업장명" STYLE="TEXT-ALIGN: Left"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtAcctCd1" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd1.value,2)"> <INPUT TYPE=TEXT NAME="txtAcctNm1" SIZE=25 tag="24">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>발생일자</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5144ma1_fpDateTime2_txtIssuedDt1.js'></script>&nbsp;~&nbsp;<script language =javascript src='./js/a5144ma1_fpDateTime2_txtIssuedDt2.js'></script></TD>

								 </TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd2" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd2.value,3)"> <INPUT TYPE=TEXT NAME="txtAcctNm2" SIZE=25 tag="24"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/a5144ma1_vspdData_vspdData.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>차변합계</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a5144ma1_OBJECT22_txtNDrAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>대변합계</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a5144ma1_OBJECT22_txtNCrAmt.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</div>
		<DIV ID="TabDiv"  SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>회계일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5144ma1_fpDateTime1_txtFromGlDt1.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/a5144ma1_fpDateTime2_txtToGlDt1.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>사업장코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtBizAreaCd.value,0)"> <INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=18 tag="24X" ALT="사업장명" STYLE="TEXT-ALIGN: Left"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="14XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd.value,2)"> <INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="24"></TD>
									<TD CLASS="TD5" NOWRAP>발생일자</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5144ma1_fpDateTime2_txtIssuedDtFR.js'></script>&nbsp;~&nbsp;<script language =javascript src='./js/a5144ma1_fpDateTime2_txtIssuedDtTO.js'></script></TD>

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
								<TD HEIGHT="100%" colspan=2 WIDTH="50%">
								<script language =javascript src='./js/a5144ma1_vspdData1_vspdData1.js'></script></TD>
								<TD HEIGHT="100%" colspan=2 WIDTH="50%">
								<script language =javascript src='./js/a5144ma1_vspdData2_vspdData2.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>차변합계</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a5144ma1_OBJECT22_txtNDrAmt1.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>대변합계</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a5144ma1_OBJECT22_txtNCrAmt1.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</div>
		<DIV ID="TabDiv"  SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>회계일</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5144ma1_fpDateTime3_txtFromGlDt2.js'></script>&nbsp;~&nbsp;
												           <script language =javascript src='./js/a5144ma1_fpDateTime4_txtToGlDt2.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>사업장코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd2" SIZE=10 MAXLENGTH=10 tag="31XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtBizAreaCd2.value,1)"> <INPUT TYPE=TEXT NAME="txtBizAreaNm2" SIZE=18 tag="34X" ALT="사업장명" STYLE="TEXT-ALIGN: Left"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtInputType2" SIZE=10 MAXLENGTH=2 tag="32XXXU" ALT="전표입력경로코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInputType2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtInputType2.value,'3')"> <INPUT TYPE="Text" NAME="txtInputTypeNm2" SIZE=18 tag="14X" ALT="전표입력경로명"></TD>
									<TD CLASS="TD5" NOWRAP>발생일자</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5144ma1_fpDateTime3_txtIssuedDt21.js'></script>&nbsp;~&nbsp;<script language =javascript src='./js/a5144ma1_fpDateTime3_txtIssuedDt22.js'></script></TD>
								 </TR>
<!--
								<TR>
									<TD CLASS="TD5" NOWRAP>모듈코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMocd" SIZE=10 MAXLENGTH=1 tag="31XXXU" ALT="모듈코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../image/btnPopup.gif" NAME="btnMocd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtMocd.value,5)"> <INPUT TYPE=TEXT NAME="txtMoNm" SIZE=18 tag="34X" ALT="모듈명" STYLE="TEXT-ALIGN: Left"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
-->
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
								<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/a5144ma1_vspdData3_vspdData3.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>차변합계</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a5144ma1_OBJECT22_txtNDrAmt2.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>대변합계</TD>
								<TD class=TD6 NOWRAP><script language =javascript src='./js/a5144ma1_OBJECT22_txtNCrAmt2.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</div>
		</TD>
	</TR>
	<TR>
		<TD <%=HGIEHT_TYPE_01%>></td>
	</TR>
	<tr>	
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="34" TABINDEX="-1">

<INPUT TYPE=hidden NAME="htxtFromGlDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtToGlDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtIssuedDt1"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtIssuedDt2"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAcctCd1"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtBizAreaCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAcctCd2"		tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="htxtFromGlDt2"		tag="44" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtToGlDt2"		tag="44" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtIssuedDt21"	tag="44" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtIssuedDt22"	tag="44" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtInputType"		tag="44" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtBizAreaCd2"	tag="44" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtMocd"			tag="44" TABINDEX="-1">


</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
 

