<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : A5124MA1
'*  4. Program Name         : 계정별 보조부 조회 
'*  5. Program Desc         : Query of Account Code
'*  6. Component List       : ADO
'*  7. Modified date(First) : 2001.11.15
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Oh, Soo Min
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/JpQuery.vbs">				</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================
Dim lgIsOpenPop
Dim IsOpenPop                                               '☜: Popup status                           
Dim lgMark                                                  '☜: 마크                                  

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "A5124MB1.asp"
Const C_MaxKey          = 20 

'========================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub


'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
	
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = 2
			
			select case Trim(.Value)

			Case  "총계" 
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(255,230,255)
			Case  "월계" 
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(230,255,255)

			End select        
	     next
	     
	     For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = 3
			select case Trim(.Value)
			Case  "일계" 
			    
			    .Col = -1 
			    .Col2 = -1
			    .BackColor = RGB(255,255,230)
			End select        
	     next
	     
    End With    
	

End Sub


'========================================================================================
Sub SetDefaultVal()
    

'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
    Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

    EndDate = "<%=GetSvrDate%>"


    Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

    StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
    EndDate   = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)

    frm1.txtFromGlDt.Text   = StartDate 
    frm1.txtToGlDt.Text     = EndDate 
	frm1.txtAmtFr.Text	= ""
	frm1.txtAmtTo.Text	= ""
    frm1.txtFromGlDt.focus

End Sub

'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("A5124MA1","S","A","V20070211",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
End Sub



'========================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
        With frm1

        .vspdData.ReDraw = False
        ggoSpread.SpreadLockWithOddEvenRowColor()
        ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
        .vspdData.ReDraw = True

        End With
    End if
End Sub



'========================================================================================
Sub InitComboBox()	
	Err.clear
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & "  order by minor_nm", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
End Sub
 


'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	'frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 0
			arrParam(0) = frm1.txtDeptCd.Alt
			arrParam(1) = "B_ACCT_DEPT A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.ORG_CHANGE_ID = " & FilterVar(parent.gChangeOrgId, "''", "S")  
			arrParam(5) = frm1.txtDeptCd.Alt
	
			arrField(0) = "A.DEPT_CD"
			arrField(1) = "A.DEPT_NM"

			arrHeader(0) = "부서코드"
			arrHeader(1) = "부서명"
		
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Select Case iWhere
		Case 0					'부서 
			frm1.txtDeptCd.focus
			frm1.txtDeptCd.value = Trim(arrRet(0))
			frm1.txtDeptNm.value = arrRet(1)
		End Select
	End If	

End Function
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBpCd(arrRet)
		lgBlnFlgChgValue = True
	End If
		
End Function
'========================================================================================
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
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBpCd(arrRet)
	End If	
	
End Function

'========================================================================================
Function SetBpCd(byval arrRet)
	frm1.txtBpCd.focus
	frm1.txtBpCd.Value    = arrRet(0)		
	frm1.txtBpNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'============================================================
'부서코드 팝업 
'============================================================
Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = parent.UCN_PROTECTED Then Exit Function
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode				'부서코드 
	arrParam(1) = frm1.txtToGlDt.Text	'날짜(Default:현재일)
'	arrParam(2) = "1"					'부서권한(lgUsrIntCd)
'	If lgIntFlgMode = parent.OPMD_UMODE then
'		arrParam(3) = "T"									' 결의일자 상태 Condition  
'	Else
'		arrParam(3) = "F"									' 결의일자 상태 Condition  
'	End If
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	End If

	frm1.txtDeptCd.focus
	frm1.txtDeptCd.value = Trim(arrRet(0))
	frm1.txtDeptNm.value = arrRet(1)
	
	call txtDeptCd_OnChange()
	
	lgBlnFlgChgValue = True
	
End Function


'----------------------------------------  OpenAcctCd()  -------------------------------------------------
'	Name : OpenAcctCd()
'	Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
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
	
	If arrRet(0) = "" Then
		Select case iWhere
		case 1
			frm1.txtBizAreaCd.focus
		case 2
			frm1.txtAcctCd.focus
		case 3
			frm1.txtAcctCd2.focus
		End select	

		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If	
	
End Function

'========================================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	Dim gPos

	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
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


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,1)
	End If
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd1.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,4)
	End If
End Function


'========================================================================================
'                       회계전표 POPUP
' ========================================================================================  
Function OpenPopupGL()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	With frm1.vspdData
	   if .maxrows > 0 Then	
		.Row = .ActiveRow
		.Col = GetKeyPos("A",14)

	
		arrParam(0) = Trim(.Text)	'회계전표번호 
		arrParam(1) = ""			'Reference번호 
	   End if	
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function


'========================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						 'Cookie Split String : CookiePage Function Use

	If Kubun = 1 Then								 'Jump로 화면을 이동할 경우 

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie "PoNo" , lsPoNo					 'Jump로 화면을 이동할때 필요한 Cookie 변수정의 
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							 'Jump로 화면이 이동해 왔을경우 

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

'========================================================================================
Sub Form_Load()
    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")

	Call InitVariables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call InitComboBox()
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Call FncSetToolBar("New")
	frm1.txtTDrAmt.allownull = False 
    frm1.txtTCrAmt.allownull = False 
    frm1.txtTSumAmt.allownull = False 
    
    frm1.txtNDrAmt.allownull = False 
    frm1.txtNCrAmt.allownull = False 
    frm1.txtNSumAmt.allownull = False 
    
    frm1.txtSDrAmt.allownull = False 
    frm1.txtSCrAmt.allownull = False 
    frm1.txtSSumAmt.allownull = False 
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================
Sub txtAcctCd_onChange()
'	If Trim(frm1.txtAcctCd.value) <> "" Then
'		Call CommonQueryRs("BAL_FG", "A_ACCT", "ACCT_CD = " & Filtervar(Trim(frm1.txtAcctCd.value), "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
'		frm1.hAcctbalfg.value = Replace(lgF0, chr(11), "")
'	Else
'		frm1.txtAcctNm.value = ""
'		frm1.hAcctbalfg.value = ""
'	End If
'	frm1.txtAcctCd2.value = ""
'	frm1.txtAcctNm2.value = ""
	
End Sub

'========================================================================================
Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtFromGlDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtFromGlDt.Focus       
    End If
End Sub

'========================================================================================
Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtToGlDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtToGlDt.Focus       
    End If
End Sub

'========================================================================================
Sub txtFromGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToGlDt.focus
	   Call FncQuery
	End If   
End Sub

'========================================================================================
Sub txtToGlDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtFromGlDt.focus
	   Call FncQuery
	End If   
End Sub

'========================================================================================
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
	ggoSpread.Source = frm1.vspdData
	Call JumpPgm()

End Sub

Function JumpPgm()
	
	Dim pvSelmvid, pvFB_fg,pvKeyVal,StrNVar,StrNPgm,pvSingle
	
	if frm1.vspddata.Maxrows  < 1 then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End if
	ggoSpread.Source = frm1.vspdData
	
	frm1.vspddata.row = 0
    frm1.vspddata.col = frm1.vspddata.Activecol

    Select case frm1.vspddata.value
    
    Case "[전표번호]"
		frm1.vspddata.row = Frm1.vspdData.ActiveRow
		if 	TRIM(frm1.vspddata.value) <> "" then
				
				pvKeyVal =   frm1.vspddata.value
				pvSingle =   frm1.txtBizAreaCd.value & chr(11) & frm1.txtBizAreaCd1.value & chr(11) & frm1.fpDateTime1.text & chr(11) & frm1.fpDateTime2.text & chr(11) 
				
				pvFB_fg = "B"
				pvSelmvid = "GL_NO"
	
					Call Jump_Pgm (	pvSelmvid, _
									pvFB_fg, _
									pvSingle,  _
									pvKeyVal)
		End if 											
	Case "[계정코드]"
		frm1.vspddata.row = Frm1.vspdData.ActiveRow
		if 	TRIM(frm1.vspddata.value) <> "" then
				
				pvKeyVal =   frm1.vspddata.value
				pvSingle  =	frm1.vspddata.value  & chr(11) & _
							frm1.txtBizAreaCd.value & chr(11) & _
							frm1.txtBizAreaCd1.value & chr(11) & _ 
							frm1.fpDateTime1.text & chr(11) & _ 
							frm1.fpDateTime2.text & chr(11)
				
				pvFB_fg = "B"
				pvSelmvid = "ACCT_CD"
	
					Call Jump_Pgm (	pvSelmvid, _
									pvFB_fg, _
									pvSingle,  _
									pvKeyVal)
		End if								
									
	Case "[거래처코드]"
		frm1.vspddata.row = Frm1.vspdData.ActiveRow
		
		if 	TRIM(frm1.vspddata.value) <> "" then
		
				pvKeyVal =   frm1.vspddata.value
				pvSingle  =	""
				pvFB_fg = "B"
				pvSelmvid = "BP_CD"
	
					Call Jump_Pgm (	pvSelmvid, _
									pvFB_fg, _
									pvSingle,  _
									pvKeyVal)										
	
			
		End if
		
		
		
		
	Case "[전표입력경로]"
		frm1.vspddata.row = Frm1.vspdData.ActiveRow
		frm1.vspddata.col = 14

		if 	TRIM(frm1.vspddata.value) <> "" then
	
				
					pvKeyVal =   frm1.vspddata.value
					
									
					pvSingle =   ""
				
					pvFB_fg = "B"
					pvSelmvid = "GL_TYPE"
	
						Call Jump_Pgm (	pvSelmvid, _
										pvFB_fg, _
										pvSingle,  _
										pvKeyVal)
										
										
										
	End if 											
		 
	End select
End Function

'========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================
Sub txtDeptCd_OnChange()
  
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	Dim lgF2By2

	If Trim(frm1.txtToGlDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtToGlDt.Text, parent.gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
'			If lgIntFlgMode <> parent.OPMD_UMODE Then
'				IntRetCD = DisplayMsgBox("124600","X","X","X")  
'			End If			
'			frm1.txtDeptCd.Value = ""
			frm1.txtDeptNm.Value = ""
			frm1.hOrgChangeId.Value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.Value = Trim(arrVal2(2))
			Next	
			
		End If
	

End Sub

'========================================================================================
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

'========================================================================================
Sub txtAmtFr_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub
'========================================================================================
Sub txtAmtTo_Keypress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call fncQuery()
    End if
End Sub

'==========================================================================================
Function CompareAcctCdByDB(ByVal FromCd , ByVal ToCd)
	Dim strSelect,strFrom,strWhere
	Dim iFlag,iRs

	CompareAcctCdByDB = False

    If FromCd.value <> "" And ToCd.value <> "" Then
        strSelect = ""
        strSelect = "  Case When  " & FilterVar(UCase(FromCd.value), "''", "S") & " "
        strSelect = strSelect & "  >  " & FilterVar(UCase(ToCd.value), "''", "S") & "  Then " & FilterVar("N", "''", "S") & "  "
        strSelect = strSelect & " When  " & FilterVar(UCase(FromCd.value), "''", "S") & " "
        strSelect = strSelect & "  <=  " & FilterVar(UCase(ToCd.value), "''", "S") & "  Then " & FilterVar("Y", "''", "S") & "  End "
        strFrom = ""
        strWhere = ""
        If CommonQueryRs2by2(strSelect, strFrom, strWhere, iRs) = True Then
            iFlag = Split(iRs, Chr(11))
            If Trim(iFlag(1)) = "N" Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    CompareAcctCdByDB = True
End Function

'==========================================================================================
Function CompareGlAmtByDB(ByVal FromAmt , ByVal ToAmt)
	Dim strSelect,strFrom,strWhere
	Dim iFlag,iRs

	CompareGlAmtByDB = False
	'문자형을 숫자형으로 바꿈 
    If FromAmt.text <> "" And ToAmt.text <> "" Then
        strSelect = ""
        strSelect = "  Case When  " & FilterVar(UNICDBL(FromAmt.text), "''", "SNM") & " "
        strSelect = strSelect & "  >  " & FilterVar(UNICDBL(ToAmt.text), "''", "SNM") & "  Then " & FilterVar("N", "''", "S") & "  "
        strSelect = strSelect & " When  " & FilterVar(UNICDBL(FromAmt.text), "''", "SNM") & " "
        strSelect = strSelect & "  <=  " & FilterVar(UNICDBL(ToAmt.text), "''", "SNM") & "  Then " & FilterVar("Y", "''", "S") & "  End "
        strFrom = ""
        strWhere = ""
        If CommonQueryRs2by2(strSelect, strFrom, strWhere, iRs) = True Then
            iFlag = Split(iRs, Chr(11))
            If Trim(iFlag(1)) = "N" Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    CompareGlAmtByDB = True
End Function

'========================================================================================
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then										'⊙: This function check indispensable field
		Exit Function
    End If

	Call txtDeptCd_OnChange()
	
	If CompareDateByFormat(frm1.txtFromGlDt.Text, frm1.txtToGlDt.Text, frm1.txtFromGlDt.Alt, frm1.txtToGlDt.Alt, _
						"970025", frm1.txtFromGlDt.UserDefinedFormat, parent.gComDateType, True) = False Then
		frm1.txtFromGlDt.focus
		Exit Function
	End If
	
    If CompareAcctCdByDB(frm1.txtAcctCd,frm1.txtAcctCd2) = False Then
        Call DisplayMsgBox("970025", "X", frm1.txtAcctCd.Alt, frm1.txtAcctCd2.Alt)
        frm1.txtAcctCd.focus
		Exit Function
	End If		

    If CompareGlAmtByDB(frm1.txtAmtFr,frm1.txtAmtTo) = False Then
        Call DisplayMsgBox("970025", "X", frm1.txtAmtFr.Alt, frm1.txtAmtTo.Alt)
        frm1.txtAmtFr.focus
		Exit Function
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If	

    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData	

	Call FncSetToolBar("New")
    Call DbQuery

    FncQuery = True		
End Function


'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function


'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub


'========================================================================================
Function FncExit()
    FncExit = True
End Function

'-------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(ByVal arrRet,ByVal field_fg) 
	With frm1	
		Select case field_fg
			case 1
				.txtBizAreaCd.focus
				.txtBizAreaCd.Value		= arrRet(0)
				.txtBizAreaNm.Value		= arrRet(1)
			case 2
				.txtAcctCd.focus
				.txtAcctCd.Value		= arrRet(0)
				.txtAcctNm.Value		= arrRet(1)
				.txtAcctCd2.Value		= arrRet(0)
				.txtAcctNm2.Value		= arrRet(1)

			case 3
				.txtAcctCd2.focus
				.txtAcctCd2.Value		= arrRet(0)
				.txtAcctNm2.Value		= arrRet(1)
'				.hAcctbalfg.Value		= arrRet(4)
			case 4
				.txtBizAreaCd1.focus
				.txtBizAreaCd1.Value		= arrRet(0)
				.txtBizAreaNm1.Value		= arrRet(1)
		End select	
	End With

End Function


'========================================================================================
Function DbQuery() 
	Dim strVal, strZeroFg

    DbQuery = False

    Err.Clear
	Call LayerShowHide(1)

    With frm1


'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtFromGlDt=" & UniConvDateToYYYYMMDD(frm1.txtFromGlDt.Text,parent.gDateFormat,"")
		strVal = strVal & "&txtToGlDt=" & UniConvDateToYYYYMMDD(frm1.txtToGlDt.Text,parent.gDateFormat,"")
		strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.Value)
		strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd.Value)
		strVal = strVal & "&txtAcctCd2=" & Trim(.txtAcctCd2.Value)
		strVal = strVal & "&hOrgChangeId=" & Trim(.hOrgChangeId.Value)
		strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.Value)
		strVal = strVal & "&txtDesc=" & Trim(.txtDesc.Value)
		strVal = strVal & "&txtProject=" & Trim(.txtProject.Value)
		strVal = strVal & "&ZeroFg=" & strZeroFg
		strVal = strVal & "&cboGlInputType=" & .cboGlInputType.value
		strVal = strVal & "&txtRefNo=" & .txtRefNo.value
		strVal = strVal & "&txtAmtFr=" & .txtAmtFr.text
		strVal = strVal & "&txtAmtTo=" & .txtAmtTo.text
		strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.value)
		strVal = strVal & "&txtBizAreaCd1=" & Trim(.txtBizAreaCd1.value)
		strVal = strVal & "&txtDeptCd_Alt=" & Trim(.txtDeptCd.Alt)
		strVal = strVal & "&txtAcctCd_Alt=" & Trim(.txtAcctCd.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt1=" & Trim(.txtBizAreaCd1.Alt)
		
		If frm1.chkFg.checked = True Then
			strVal = strVal & "&txtFG=" & "2"
		Else
			strVal = strVal & "&txtFG=" & "1"
		End If
		

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------

		strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		'msgbox strval
        Call RunMyBizASP(MyBizASP, strVal)

    End With

    DbQuery = True

End Function



'========================================================================================
Function DbQueryOk()
'    Call ggoOper.LockField(Document, "Q")

	IF Trim(frm1.txtdeptcd.value) = "" then
		frm1.txtdeptnm.value = ""
	end if	
	Call FncSetToolBar("Query")
	CALL InitData()
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정별보조부조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>회계일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="시작회계일자" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
												           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="종료회계일자" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDeptCd" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopupDept(frm1.txtDeptCd.value, 0)">
									                       <INPUT TYPE=TEXT NAME="txtDeptNm" ALT="부서명" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="시작계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd.value,2)"> <INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="14">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd2" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="종료계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd2.value,3)"> <INPUT TYPE=TEXT NAME="txtAcctNm2" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd1()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=30 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAmtFr" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="시작금액" id=OBJECT1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAmtTo" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="11XXXX" ALT="종료금액" id=OBJECT2></OBJECT>');</SCRIPT>
										 </TD>	
									<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="11" STYLE="WIDTH:82px:" ALT="전표입력경로"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
								 </TR>
								 <TR>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=12 MAXLENGTH=10 tag="11XXXU" ALT="거래처코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:call OpenBp(frm1.txtbpcd.value,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="24" ALT="거래처"></TD>																		
									<TD CLASS=TD5 NOWRAP>참조번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호" MAXLENGTH="30" SIZE="20" tag="11XXXU" ></TD></TD>				
								 </TR>
								 <TR>
									<TD CLASS=TD5 NOWRAP>적요</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDesc" ALT="적요" MAXLENGTH="128" SIZE="50" tag="11" ></TD>
									<TD CLASS=TD5 NOWRAP>프로젝트번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtProject" ALT="프로젝트번호" MAXLENGTH="25" SIZE="25" tag="11" >
								 	<LABEL FOR=chkConfFg>조회구분</LABEL>
									<INPUT type="checkbox" CLASS="STYLE CHECK"  NAME=chkFg ID=chkFg tag="1"></TD>
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
								<TD HEIGHT="100%" colspan=7>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>이월금액</TD>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="차변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="대변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="잔액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발생금액</TD>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="차변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="대변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>차이</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="차이" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>누계금액</TD>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="차변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="대변" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD class=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="잔액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
							</TR>
						</TABLE>						
					</TD>
				</TR>
			</TABLE>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1">
</TEXTAREA><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="34" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
 

