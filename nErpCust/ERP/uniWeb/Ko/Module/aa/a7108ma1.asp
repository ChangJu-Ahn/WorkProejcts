<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Change
'*  3. Program ID           : a7107ma1
'*  4. Program Name         : 고정자산 매각/폐기내역 등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : AS0031,  
'							  AS0039	
'							  +B19029LookupNumericFormatF	
'*  7. Modified date(First) : 2000/03/18
'*  8. Modified date(Last)  : 2001/06/02
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'=======================================================================================================
'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit    												'☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'@PGM_ID
Const BIZ_PGM_ID  = "a7108mb1.asp"  
											'비지니스 로직 ASP명 
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'환율정보 비지니스 로직 ASP명 

'@Grid_Column
Dim C_Seq	
Dim C_RcptType							            'Spread Sheet 의 Columns 인덱스 
Dim C_RcptTypePopup
Dim C_RcptTypeNm								            'Spread Sheet 의 Columns 인덱스 
Dim C_Amt
Dim C_LocAmt
Dim C_BankAcct
Dim C_BankAcctPopup
Dim C_NoteNo
Dim C_NoteNoPopup

Const C_SHEETMAXROWS = 30							            '한 화면에 보여지는 최대갯수 

'@Global_Var
'Dim lgBlnFlgChgValue                                         'Variable is for Dirty flag
'Dim lgIntGrpCount                                            'Group View Size를 조사할 변수 
'Dim lgIntFlgMode                                             'Variable is for Operation Status

'Dim lgStrPrevKey                                             'Previous NextKey

Dim IsOpenPop						                        'Popup
'Dim lgSortKey
Dim gSelframeFlg                                            'Current Tab Page

Dim lgMasterQueryFg                                         ''자산Master의 query 여부 

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'======================================================================================================

Sub initSpreadPosVariables()
	C_Seq				= 1
	C_RcptType			= 2									            'Spread Sheet 의 Columns 인덱스 
	C_RcptTypePopup		= 3
	C_RcptTypeNm		= 4								            'Spread Sheet 의 Columns 인덱스 
	C_Amt				= 5
	C_LocAmt			= 6
	C_BankAcct			= 7
	C_BankAcctPopup		= 8
	C_NoteNo			= 9
	C_NoteNoPopup		= 10
End Sub


'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                                           'initializes Group View Size
	lgBlnFlgChgValue = False
    lgStrPrevKey = ""                                           'initializes Previous Key
	
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
<%
	Dim svrDate
	svrDate = GetSvrDate
%>

	frm1.txtChgDt.text    = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
	frm1.txtDueDt.text     = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
	frm1.txtIssuedDt.text    = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
	frm1.txtDocCur.value	= parent.gCurrency
	frm1.txtXchRate.text	= "1"

	lgBlnFlgChgValue = False
End Sub



'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>  ' check
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub



'=======================================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
		.ReDraw = false	
		
		.MaxCols = C_NoteNoPopup + 1                               '☜: 최대 Columns의 항상 1개 증가시킴 
		'.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear
		
		'Hidden Column 설정 
    	.Col = .MaxCols											'공통콘트롤 사용 Hidden Column
    	.ColHidden = True
    		
'    	.Col = C_RcptType
'    	.ColHidden = True

		Call GetSpreadColumnPos("A")
		
'		Call AppendNumberPlace("6","3","0")

		'ggoSpread.SSSetFloat  C_Seq,       "순번",       5, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit	  C_Seq,        "순번",       5, 2, -1, 5
		ggoSpread.SSSetEdit  C_RcptType,  "지급유형"       ,10,,,5,2           
		ggoSpread.SSSetButton C_RcptTypePopup
		ggoSpread.SSSetEdit  C_RcptTypeNm,"지급유형명"     ,16

		ggoSpread.SSSetFloat  C_Amt,       "금액",       19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_LocAmt,    "금액(자국)", 19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		
		ggoSpread.SSSetEdit	  C_BankAcct,  "예적금코드",   25, 0, -1, 30,2
		ggoSpread.SSSetButton C_BankAcctPopup
		ggoSpread.SSSetEdit   C_NoteNo,    "어음번호",     25, 0, -1, 30,2
		ggoSpread.SSSetButton C_NoteNoPopup
		
		Call ggoSpread.MakePairsColumn(C_RcptType,C_RcptTypePopup,"1")
		Call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPopup,"1")
		Call ggoSpread.MakePairsColumn(C_NoteNo,C_NoteNoPopup,"1")
		
		.ReDraw = true
		
		Call SetSpreadLock 
		
	End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadLock C_Seq,        -1, C_Seq
	    ggoSpread.SpreadLock C_NoteNoPopup+1,   -1, C_NoteNoPopup+1

		
		.vspdData.ReDraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
   	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_Seq, pvStarRow, pvEndRow	
		ggoSpread.SSSetRequired	 C_RcptType, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_RcptTypeNm, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	 C_Amt, pvStarRow, pvEndRow
		
		.vspdData.ReDraw = True
	End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_Seq				= iCurColumnPos(1)
		C_RcptType			= iCurColumnPos(2)
		C_RcptTypePopup		= iCurColumnPos(3)
		C_RcptTypeNm		= iCurColumnPos(4)
		C_Amt				= iCurColumnPos(5)
		C_LocAmt			= iCurColumnPos(6)
		C_BankAcct			= iCurColumnPos(7)
		C_BankAcctPopup		= iCurColumnPos(8)
		C_NoteNo			= iCurColumnPos(9)
		C_NoteNoPopup		= iCurColumnPos(10)
	End Select
End Sub

'======================================================================================================
'   Function Name : OpenChgNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenChgNoInfo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	iCalledAspName = AskPRAspName("a7107ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7107ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName & "?PID=" & gstrRequestMenuID , Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtChgNo.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetChgNoInfo(arrRet)
	End If	

	
End Function

'======================================================================================================
'   Function Name : SetChgNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetChgNoInfo(Byval arrRet)
	With frm1
		.txtChgNo.value  = Trim(arrRet(0))
	End With
End Function

'=======================================================================================================
'	Name : OpenDeptCd()
'	Description : Dept Cd PopUp
'=======================================================================================================
Function OpenDeptCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strAsstNo
	Dim IntRetCd

	If IsOpenPop = True  Then Exit Function

	strAsstNo  = Trim(frm1.txtAsstNo.value)
	
	if strAsstNo = "" then
		IntRetCD = DisplayMsgBox("117326","X","X","X")    '자산번호를 입력하십시오.
		IsOpenPop = False
		Exit Function
	end if
			
			arrParam(0) = "회계부서 팝업"	
			arrParam(1) = "B_ACCT_DEPT"
			arrParam(2) = Trim(frm1.txtDeptCd.value)
			arrParam(3) = ""
			arrParam(4) = " INTERNAL_CD IN (SELECT INTERNAL_CD FROM A_ASSET_INFORM_OF_DEPT WHERE ASST_NO =  " & FilterVar(frm1.txtAsstNo.value, "''", "S") & " )"
			arrParam(5) = "지출부서"

			arrField(0) = "DEPT_CD"	
			arrField(1) = "DEPT_NM"
			arrField(2) = "ORG_CHANGE_ID "

			arrHeader(0) = "회계부서코드"
			arrHeader(1) = "회계부서명"
			arrHeader(2) = "조직개편ID"

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtDeptCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDeptCd(arrRet)
	End If
End Function


'=======================================================================================================
'	Name : SetDeptCd()
'	Description : DeptCd Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetDeptCd(byval arrRet)
	frm1.txtDeptCd.Value		= Trim(arrRet(0))
	frm1.txtDeptNm.Value		= arrRet(1)
	frm1.HORGCHANGEID.Value		= Trim(arrRet(2))
	lgBlnFlgChgValue = True
End Function
Function OpenDept()

	Dim arrRet
	Dim arrParam(8)

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function

	arrParam(0) = Trim(frm1.txtDeptCd.value) 'strCode		            '  Code Condition
   	arrParam(1) = frm1.txtChgDt.Text

	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,"DeptCd")
	End If	
End Function

Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
		
			case "DeptCd"
				.txtChgDt.text			= arrRet(3)
				.txtDeptCd.value        = arrRet(0)
				.txtDeptNm.value 		= arrRet(1)
				Call txtDeptCd_OnChange()

		End select	

		lgBlnFlgChgValue = True
	End With
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	if frm1.Rb_Duse.checked = True then Exit Function   ''폐기일 때,거래처 선택 못하도록.
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
Function SetBpCd(byval arrRet)
	frm1.txtBpCd.focus
	frm1.txtBpCd.Value    = arrRet(0)		
	frm1.txtBpNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function



'=======================================================================================================
'	Name : OpenNoteNo()
'	Description : Note No PopUp
'=======================================================================================================
Function OpenNoteNo(Byval strCode,Byval strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function

	IF UCase(strCard) = "CR" Then
		arrParam(0) = "수취구매카드 팝업"				        ' 팝업 명칭 
		arrParam(1) = "f_note a,b_biz_partner b, b_bank c, b_card_co d"		' TABLE 명칭 
		arrParam(2) = ""								' Code Condition
		arrParam(3) = ""								' Name Cindition
		arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CR", "''", "S") & "  and a.bp_cd = b.bp_cd  "
		arrParam(4) = arrParam(4) & " and a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "
		arrParam(5) = "구매카드번호"						' 조건필드의 라벨 명칭 

		arrField(0) = "a.Note_no"					' Field명(0)
		arrField(1) = "F2" & parent.gColSep & "a.Note_amt"		' Field명(1)
		arrField(2) = "DD" & parent.gColSep & "a.Issue_dt"		' Field명(2)
		arrField(3) = "b.bp_nm"					' Field명(3)
		arrField(4) = "d.card_co_nm"    	    			' Field명(4)

		arrHeader(0) = "구매카드번호"				' Header명(0)
		arrHeader(1) = "금액"				' Header명(1)
		arrHeader(2) = "발행일"				' Header명(2)
		arrHeader(3) = "거래처"				' Header명(3)
		arrHeader(4) = "카드사"				' Header명(4)

	Else

		arrParam(0) = "어음번호 팝업"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"
		arrParam(2) = strCode
		arrParam(3) = ""

		arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"
		arrParam(5) = "어음번호"

		arrField(0) = "A.NOTE_NO"
		arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
		arrField(2) = "C.BP_NM"
		arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
		arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"
		arrField(5) = "B.BANK_NM"

		arrHeader(0) = "어음번호"
		arrHeader(1) = "어음금액"
		arrHeader(2) = "거래처"
		arrHeader(3) = "발행일"
		arrHeader(4) = "만기일"
		arrHeader(5) = "은행"
	End if

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	Call SetActiveCell(frm1.vspdData,C_NoteNo,frm1.vspdData.ActiveRow ,"M","X","X")
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetNoteNo(arrRet)
	End If
End Function

'=======================================================================================================
'	Name : SetNoteNo()
'	Description : Note No Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetNoteNo(byval arrRet)
	With frm1

		.vspdData.Col	= C_NoteNo
		.vspdData.Text	= arrRet(0)

		.vspdData.Col	= C_Amt
		.vspdData.Text	= arrRet(1)

		.vspdData.Col	= C_LocAmt
		.vspdData.Text	= arrRet(1)

	    Call vspdData_Change(.vspdData.Col, .vspdData.Row)				 ' 변경이 dlf어났다고 알려줌 
		lgBlnFlgChgValue = True
	End With
End Function

'=======================================================================================================
'	Name : OpenCurrency()
'	Description : Currency PopUp
'=======================================================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

    if frm1.Rb_Duse.checked = True then Exit Function
	If IsOpenPop = True Then Exit Function

	arrParam(0) = "거래통화 팝업"	
	arrParam(1) = "B_CURRENCY"
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래통화"

    arrField(0) = "CURRENCY"
    arrField(1) = "CURRENCY_DESC"

    arrHeader(0) = "거래통화"
    arrHeader(1) = "거래통화명"

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtDocCur.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetCurrency()
'	Description : Currency Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCurrency(byval arrRet)
	frm1.txtDocCur.value    = arrRet(0)
	lgBlnFlgChgValue = True
	If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' 거래통화하고 Company 통화가 다를때 환율을 0으로 셋팅 
		frm1.txtXchRate.text	= "0"
	Else
		frm1.txtXchRate.text	= "1"
	End If

	call txtDocCur_OnChangeASP()
End Function

'=======================================================================================================
'	Name : OpenBankAcct()
'	Description : Bank Account No PopUp
'=======================================================================================================
Function OpenBankAcct(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function	

	arrParam(0) = "예적금코드 팝업"	' 팝업 명칭 
	arrParam(1) = "B_BANK A, F_DPST B"			' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "A.BANK_CD = B.BANK_CD "		' Where Condition
	arrParam(5) = "은행코드"				' 조건필드의 라벨 명칭 

	arrField(0) = "A.BANK_NM"					' Field명(1)
	arrField(1) = "B.BANK_ACCT_NO"				' Field명(2)

	arrHeader(0) = "은행명"						' Header명(1)
	arrHeader(1) = "예적금코드"

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	Call SetActiveCell(frm1.vspdData,C_BankAcct,frm1.vspdData.ActiveRow ,"M","X","X")
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBankAcct(arrRet)
	End If	

End Function

'=======================================================================================================
'	Name : SetBankAcct()
'	Description : Bank Account No Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetBankAcct(byval arrRet)
	With frm1
		.vspdData.Col = C_BankAcct
		.vspdData.Text = arrRet(1)
	    Call vspdData_Change(.vspdData.Col, .vspdData.Row)				 ' 변경이 읽어났다고 알려줌 
'		lgBlnFlgChgValue = True
	End With
End Function

 '------------------------------------------  OpenPoRef()  -------------------------------------------------
'	Name : OpenPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenMasterRef()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	If lgIntFlgMode = parent.OPMD_UMODE Then 
			Call DisplayMsgBox("200005", "X", "X", "X")
			Exit function
	End If	
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName & "?PID=" & gstrRequestMenuID, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtDeptCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPoRef(arrRet)
	End If		
		
End Function
 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)

    Dim strVal

    lgMasterQueryFg = False


	Call ggoOper.ClearField(Document, "A")
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    Call SetDefaultVal

    Call InitVariables

	frm1.txtAsstNo.value     = strRet(0)
	frm1.txtAsstNm.value	 = strRet(1)
	frm1.txtRegDt.text       = strRet(2)

	frm1.txtAcctDeptNm.value = strRet(9)

	frm1.txtAcqQty.text     = strRet(7)
	frm1.txtInvQty.text     = strRet(8)
	frm1.txtChgQty.text		= strRet(8) 'jsk 2003/09/23 매각폐기수량 

	frm1.txtCur.value		 = strRet(3)
	frm1.txtXchRt.text       = strRet(4)

	frm1.txtAcqAmt.text     = strRet(5)
	frm1.txtAcqLocAmt.text  = strRet(6)
	frm1.txtChgQty.text  = strRet(8)
	

	frm1.txtDeptCd.focus

'	Call ggoOper.LockField(Document, "Q")

	lgMasterQueryFg = True

	IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
	call txtCur_OnChange()
	'''lgBlnFlgChgValue = False
	End If
End Sub


'===========================================================================
' Function Name : OpenReportAreaCd
' Function Desc : OpenReportAreaCd Reference Popup
'===========================================================================
Function OpenReportAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	If frm1.txtReportAreaCd.className = parent.UCN_PROTECTED Then Exit Function

	arrParam(0) = "신고사업장 팝업"
	arrParam(1) = "B_TAX_BIZ_AREA"
	arrParam(2) = Trim(frm1.txtReportAreaCd.value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "신고사업장"

    arrField(0) = "TAX_BIZ_AREA_CD"
    arrField(1) = "TAX_BIZ_AREA_NM"

    arrHeader(0) = "신고사업장코드"
    arrHeader(1) = "신고사업장명"

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtReportAreaCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReportArea(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetReportArea()
'	Description : Bp Cd Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetReportArea(byval arrRet)
	frm1.txtReportAreaCd.Value		= arrRet(0)
	frm1.txtReportAreaNm.Value		= arrRet(1)
	lgBlnFlgChgValue = True
End Function



'===========================================================================
' Function Name : OpenArAcct()
' Function Desc : OpenArAcct Reference Popup
'===========================================================================
Function OpenArAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "미수금계정 팝업"
	arrParam(1) = "a_jnl_acct_assn a, a_acct b"
	arrParam(2) = Trim(frm1.txtArAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.trans_type = " & FilterVar("AS003", "''", "S") & "  and A.Acct_cd = B.Acct_cd and Jnl_cd = " & FilterVar("AR", "''", "S") & " "
	arrParam(5) = "미수금계정 코드"

    arrField(0) = "a.acct_cd"
    arrField(1) = "b.acct_nm"

    arrHeader(0) = "미수금계정 코드"
    arrHeader(1) = "미수금계정명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtArAcctCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetArAcct(arrRet)
	End If
End Function

'=======================================================================================================
'	Name : SetArAcct()
'	Description : 
'=======================================================================================================
Function SetArAcct(byval arrRet)
	frm1.txtArAcctCd.Value		= arrRet(0)
	frm1.txtArAcctNm.Value		= arrRet(1)
	lgBlnFlgChgValue = True
End Function

Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrParamAdo(3)

	If IsOpenPop = True Then Exit Function	
	
	Select Case iWhere
		Case 6    
			arrParam(0) = "입금유형"								' 팝업 명칭 
		 
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = Trim(frm1.vspdData.text)
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "(A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " ) " _
			   & " AND A.MINOR_CD NOT IN ( " & FilterVar("NP", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("AP", "''", "S") & " , " & FilterVar("CP", "''", "S") & "  , " & FilterVar("NE", "''", "S") & " , " & FilterVar("PR", "''", "S") & " ) AND B.SEQ_NO = 4 " ' Where Condition        
			arrParam(5) = "입금유형"								' TextBox 명칭 
	 
			arrField(0) = "A.MINOR_CD"							' Field명(0)
			arrField(1) = "A.MINOR_NM"							' Field명(1)
			  
			arrHeader(0) = "입금유형"								' Header명(0)
			arrHeader(1) = "입금유형명"								' Header명(1) 
	End Select
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			

	IsOpenPop = False

	Call GridSetFocus(iWhere)
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If

End Function
'=======================================================================================================
Function GridsetFocus(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 6
				Call SetActiveCell(.vspdData,C_Rcpttype,.vspdData.ActiveRow ,"M","X","X")
		END Select
	End With
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
	Select Case iWhere
		Case 6
			.vspdData.Col = C_RcptType
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_RcptTypeNm
			.vspdData.Text = arrRet(1)
			Call vspdData_Change(C_RcptType, frm1.vspdData.Row)				 ' 변경이 읽어났다고 알려줌 		
	END Select
	End With
	lgBlnFlgChgValue = True
End Function



Function OpenVatType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function	
	if frm1.Rb_Duse.checked = True then Exit Function

	arrHeader(0) = "부가세유형"						' Header명(0)
	arrHeader(1) = "부가세명"						' Header명(1)
	arrHeader(2) = "부가세Rate"

	arrField(0) = "B_Minor.MINOR_CD"							' Field명(0)
	arrField(1) = "B_Minor.MINOR_NM"							' Field명(1)
    arrField(2) = "F2" & parent.gColSep & "b_configuration.REFERENCE"
'	arrField(2) = "b_configuration.REFERENCE"

	arrParam(0) = "부가세유형"						' 팝업 명칭 
	arrParam(1) = "B_Minor,b_configuration"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtVatType.value)			' Code Condition
	'arrParam(3) =	""		' Name Cindition

	arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
	              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd = B_Minor.Major_Cd"
	arrParam(5) = "부가세유형"						' TextBox 명칭 

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtVatType.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetVatType(arrRet)
	End If
End Function

'=======================================================================================================
'	Name : Setvattype()
'	Description : Bp Cd Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetVatType(byval arrRet)
	frm1.txtVatType.Value    = arrRet(0)
	frm1.txtVatTypeNm.Value    = arrRet(1)
	frm1.txtVatRate.text    = arrRet(2)
	Call txtVatType_OnChange
	lgBlnFlgChgValue = True
End Function

'======================================================================================================
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'=======================================================================================================

Function OpenPopupTempGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function
'=======================================================================================================
'Description : 회계전표 생성내역 팝업 
'=======================================================================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function


Function MaxSpreadVal(byval Row)
  Dim iRows
  Dim MaxValue  
  Dim tmpVal

	MAxValue = 0
	with frm1
		For iRows = 1 to  .vspdData.MaxRows
			.vspddata.row = iRows
		        .vspddata.col = C_Seq

			if .vspdData.Text = "" then 
			   tmpVal = 0
			else
  			   tmpVal = cdbl(.vspdData.value)
			end if

			if tmpval > MaxValue   then
			   MaxValue = cdbl(tmpVal)

			end if
		Next

		MaxValue = MaxValue + 1

		.vspddata.row = row
		.vspddata.col = C_Seq
		.vspddata.text = MaxValue
	end with

end Function
 
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

   ' ------ Developer Coding part (Start ) --------------------------------------------------------------
   Dim strCodeList
    Dim strNameList
	'jsk 20030828 NR- > NP 
	Call CommonQueryRs("A.MINOR_CD,A.MINOR_NM","B_MINOR A, B_CONFIGURATION B", _
						"(A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " ) AND A.MINOR_CD NOT IN ( " & FilterVar("NP", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("AP", "''", "S") & " , " & FilterVar("CP", "''", "S") & "  , " & FilterVar("NE", "''", "S") & " , " & FilterVar("PR", "''", "S") & " ) AND B.SEQ_NO = 4 ", _
	                         lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	'A1006

    strCodeList = Replace(lgF0, Chr(11), vbTab)
    strNameList = Replace(lgF1, Chr(11), vbTab)

    ggoSpread.SetCombo strCodeList, C_RcptType
    ggoSpread.SetCombo strNameList, C_RcptTypeNm

    '------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	Dim iRow
	Dim varData
	
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
	Call SetSpreadColor(-1,-1)

	With frm1					
		.vspdData.Redraw = False
		For iRow = 1 To frm1.vspdData.MaxRows
			.vspdData.Col = C_RcptType		
			.vspdData.Row = iRow
			varData = .vspdData.text
			
			Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type별 입력필수 필드 표시 
		Next
		.vspdData.Redraw = True			
	End With
	
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()
'    Call GetGlobalVar
'    Call ClassLoad                                                          'Load Common DLL
    Call LoadInfTB19029                                                     'Load table , B_numeric_format
        
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
                                                                            'Format Numeric Contents Field                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    frm1.txtAcqAmt.AllowNull =false
    frm1.txtTotalAmt.AllowNull =false
    frm1.txtApAmt.AllowNull =false
    frm1.txtDeprAmt.AllowNull =false
        
    Call InitSpreadSheet                                                    'Setup the Spread sheet
    Call InitVariables                                                      'Initializes local global variables
    Call SetDefaultVal
    frm1.hORGCHANGEID.value =parent.gChangeOrgId 
    
    Call SetToolBar("1110110100100111")										' 처음 로드시 표준 에 따라 

	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
	
    frm1.txtChgNo.focus 

End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtChgDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtChgDt.Focus
    End If
End Sub
Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDueDt.Focus
    End If
End Sub

Sub txtIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtIssuedDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_Change()
'   Event Desc : 
'=======================================================================================================
Sub txtChgDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtIssuedDt_Change()
    lgBlnFlgChgValue = True
End Sub



'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )
Dim DblEntryQty

    With frm1.vspdData 
        If Col = C_Amt then
            .Col = C_Amt
            If .Text = "" Then
                DblEntryQty = 0
            Else
                DblEntryQty = UNICDbl(.Text)
            End If
        End If
    
    End With
                
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim intIndex
	Dim varData

	With frm1.vspdData
	
		.Row = Row
    
		frm1.vspdData.ReDraw = False
		Select Case Col
			Case  C_RcptType
				.Col = Col
				intIndex = .Value
				.Col = C_RcptType
				.Value = intIndex
				varData = .text
				If Trim(varData) <> "" Then 
					IF CommonQueryRs( " A.MINOR_CD,A.MINOR_NM " , "B_MINOR A, B_CONFIGURATION B  " , "  (A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						Select Case UCase(lgF0)
							Case "DP" & Chr(11)			' 예적금 
								.Row  = Row
								.Col  = C_NoteNo
								.Text = ""
							Case "NO" & Chr(11)
								.Row  = Row
								.Col  = C_BankAcct
								.Text = ""
							Case Else
								.Row  = Row
								.Col  = C_NoteNo
								.Text = ""

								.Row  = Row
								.Col  = C_BankAcct
								.Text = ""
						End Select
						.Col  = C_RcptTypeNm
						.Text = Replace(lgF1, Chr(11), "")
					Else
						Call DisplayMsgBox("179051", "X", "X" ,"x")
						.Col  = C_RcptType
						.Text = ""
						.Col  = C_RcptTypeNm
						.Text = ""
						Call SetActiveCell(frm1.vspdData,C_RcptType,frm1.vspdData.ActiveRow ,"M","X","X")
					End if
				End if

				'.Col  = C_Amt
				'.Text = ""
				'.Col  = C_LocAmt
				'.Text = ""

				call subVspdSettingChange(Row,varData)
		End Select
	End With

	frm1.vspdData.ReDraw = True	
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row   

    lgBlnFlgChgValue = True

End Sub


'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    Dim iColumnName
    
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if

End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) 컬럼 width 변경 이벤트 핸들러 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub


'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' 거래통화하고 Company 통화가 다를때 환율을 0으로 셋팅 
		frm1.txtXchRate.text	= "0"                         ' 디폴트값인 1이 들어가 있으면 환율이 입력된 것으로 판단하여 
								                                        ' 환율정보를 읽지 않고 입력된 값으로 계산. 
	Else 

		frm1.txtXchRate.text	= "1"
	End If	
    
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub

'==========================================================================================
'   Event Name : txtVatType_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtVatType_OnChange()
	Dim dblVatAmt
	Dim StrType, StrName, StrRate
	
	lgBlnFlgChgValue = True

	Call CommonQueryRs(" A.MINOR_CD, A.MINOR_NM, B.REFERENCE ", " B_MINOR A JOIN B_CONFIGURATION B ON A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD ", _
		" A.MAJOR_CD = " & FilterVar("B9001", "''", "S") & "  AND	B.SEQ_NO = 1 " & " AND A.MINOR_CD =  " & FilterVar(frm1.txtVatType.value , "''", "S") & "" , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	
	StrType = Split(lgF0, Chr(11))
	StrName = Split(lgF1, Chr(11))
	StrRate = Split(lgF2, Chr(11))
	
	If Trim(lgF0) <> "" then
		frm1.txtVatType.value = StrType(0)
		frm1.txtVatTypeNm.value = StrName(0)
		frm1.txtVatRate.text = CDBL(StrRate(0))
	end if 
	
	
	if frm1.txtVatAmt.text = "" then
		dblVatAmt = 0
	else
		dblVatAmt = UNICDbl(frm1.txtVatAmt.text)	
	end if
	
	If Trim(frm1.txtVatType.Value) = "" and dblVatAmt = 0 Then
		ggoOper.SetReqAttr frm1.txtVatAmt, "D"    '부가세금액 
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '부가세타입 
	Else
		ggoOper.SetReqAttr frm1.txtVatAmt, "N"    '부가세금액 
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '부가세타입 
	End If

End Sub


'==========================================================================================
'   Event Name : txtVatAmt_Change
'   Event Desc : 
'==========================================================================================
Sub txtVatAmt_Change()
	Dim dblVatAmt

	lgBlnFlgChgValue = True	
	
	if frm1.txtVatAmt.text="" then
		dblVatAmt = 0
	else
		dblVatAmt = UNICDbl(frm1.txtVatAmt.text)	
	end if
		
	If dblVatAmt = 0 and Trim(frm1.txtVatType.Value) = "" Then
		ggoOper.SetReqAttr frm1.txtVatAmt, "D"    '부가세금액 
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '부가세타입 
	Else
		ggoOper.SetReqAttr frm1.txtVatAmt, "N"    '부가세금액 
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '부가세타입 
	End IF
		
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChangeASP()
    
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub

'==========================================================================================
'   Event Name : txtCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtCur_OnChange()

    lgBlnFlgChgValue = True
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCXRef()
	END IF	    
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub subVspdSettingChange(ByVal lRow, Byval varData)	
	ggoSpread.Source = frm1.vspdData
		
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)				
				Case "DP" & Chr(11)			' 예적금 
					ggoSpread.SSSetRequired	 C_BankAcct,		 lRow, lRow			
					ggoSpread.SpreadUnLock   C_BankAcct,      lRow, C_BankAcct
					ggoSpread.SpreadUnLock   C_BankAcctPopUp, lRow, C_BankAcctPopUp

					'ggoSpread.SSSetEdit		 C_BankAcct, "예적금코드", 25, 0, lRow, 30,2  
		
					ggoSpread.SSSetRequired	 C_BankAcct,      lRow, lRow	
												
					ggoSpread.SpreadLock     C_NoteNo,		 lRow, C_NoteNo,lRow   '어음번호 protect
					ggoSpread.SSSetProtected C_NoteNo,       lRow, lRow						
					ggoSpread.SpreadLock     C_NoteNoPopup,  lRow, C_NoteNoPopup,lRow          	
				Case "NO" & Chr(11)				
					ggoSpread.SpreadUnLock   C_NoteNo,        lRow, C_NoteNo,       lRow
					ggoSpread.SpreadUnLock   C_NoteNoPopup,   lRow, C_NoteNoPopup,  lRow
					 
					ggoSpread.SpreadLock     C_BankAcct,      lRow, C_BankAcct,     lRow   
					ggoSpread.SpreadLock     C_BankAcctPopup, lRow, C_BankAcctPopup,lRow
		
					ggoSpread.SSSetProtected C_BankAcct,      lRow, lRow								
		
					'ggoSpread.SSSetEdit      C_NoteNo, "어음번호", 25, 0, lRow, 30,2
					ggoSpread.SSSetRequired  C_NoteNo,        lRow, lRow
		
				Case Else									
					ggoSpread.SpreadLock     C_BankAcct,      lRow, C_BankAcct,     lRow   			
					ggoSpread.SpreadLock     C_BankAcctPopup, lRow, C_BankAcctPopup,lRow
							
					ggoSpread.SSSetProtected C_BankAcct,      lRow, lRow							
		
					ggoSpread.SpreadLock     C_NoteNo,        lRow, C_NoteNo,     lRow
					ggoSpread.SpreadLock     C_NoteNoPopup,   lRow, C_NoteNoPopup,lRow		
		
					ggoSpread.SSSetProtected C_NoteNo,        lRow, lRow													
			End Select			
		
	End if
	
End Sub	



'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
    Dim strTemp
    Dim intPos1
    Dim bankCode
	Dim intRetCd
	Dim strData
	Dim strCard
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		if Row > 0 And Col = C_BankAcctPopup Then

			.Col = C_BankAcct
			.Row = Row
			
			Call OpenBankAcct(.Text)
		
		Elseif Row > 0 And Col = C_NoteNoPopUp Then
			
'			.Col   = C_NoteNo
'			strData = .Text
			.Col = C_NoteNo
			.Row = Row
			strTemp = Trim(.text)				    
			.col = C_RcptType
			strCard = .text
			
			Call OpenNoteNo(strData, strCard)
		Elseif Row > 0 And Col = C_RcptTypePopup Then
			.Col = C_RcptType
			.Row = Row
			Call OpenPopup(.Text, 6)
		End If
	
	End With
End Sub


Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

		If Row >= NewRow Then
		    Exit Sub
		End If

    End With

End Sub

Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
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
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
    End if
        
End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  C_RcptTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_RcptType
				.Value = intIndex
				varData = .text
		End Select
	End With	

	frm1.vspdData.ReDraw = False		
	
	 IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)					
				Case "DP" & Chr(11)			' 예적금 
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_NoteNo
					frm1.vspdData.Text = ""
		
				Case "NO" & Chr(11)											
				
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_BankAcct
					frm1.vspdData.Text = ""			
				Case Else
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_NoteNo
					frm1.vspdData.Text = ""			
							
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_BankAcct
					frm1.vspdData.Text = ""				
			End Select			
	end if

	call subVspdSettingChange(Row,varData)

	frm1.vspdData.ReDraw = True	

End Sub


'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables                                                      'Initializes local global variables
'    Call InitSpreadSheet																			'⊙: Initializes local global variables
	'frm1.vspdData.MaxRows = 0 ' InitSpreadSheet 대신 
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    'Call InitComboBox

  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery															'Query db data
       
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                                          
	
	'-----------------------
	'Check previous data area
	'-----------------------
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------

	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
	Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
	Call InitVariables                                                      'Initializes local global variables

	Call SetDefaultVal
    Call SetToolBar("1110110100100111")										' 처음 로드시 표준 에 따라 

    call txtDocCur_OnChangeASP()  

	lgBlnFlgChgValue = False	
	if frm1.Rb_Duse.checked = True then    '매각일 때,	
		call Radio2_onChange()
	end if

	FncNew = True 

End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================

Function FncDelete() 
    Dim IntRetCD
	FncDelete = False
		
	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")   '삭제하시겠습니까?  
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	'-----------------------
	'Precheck area
	'-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")                                
    	Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete                                                          '☜: Delete db data
    
    FncDelete = True

End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim lDelRows, intRows
	FncSave = False

	Err.Clear

    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer   

    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then               '⊙: Check required field(Single area)
       Exit Function
    End If

	if isNull(frm1.txtApAmt.text) then
		frm1.txtApAmt.text = "0"
	end if

	if frm1.txtApAmt.text = "" then
		frm1.txtApAmt.text = "0"
	end if			

'	if IsNull(frm1.txtIssuedDt.text) or Trim(frm1.txtIssuedDt.text) = "" then
'		frm1.txtIssuedDt.text = frm1.txtChgDt.text		
'	end if
	
	If CompareDateByFormat(frm1.txtRegDt.text,frm1.txtChgDt.text,frm1.txtRegDt.Alt,frm1.txtChgDt.Alt, _
        	               "970023",frm1.txtRegDt.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtChgDt.focus
	   Exit Function
	End If
	
	if frm1.Rb_Sold.checked = True then    '매각일 때,
		'if frm1.vspdData.MaxRows < 1 then  
			'if  frm1.txtApAmt.value = 0 then
			'	IntRetCD = DisplayMsgBox("117991","X","X","X")  ''자산지출 금액을 입력하십시오.
			'	Exit Function
			'end if		
		'end if
	
		ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
		If Not ggoSpread.SSDefaultCheck         Then              '⊙: Check required field(Multi area)
		   Exit Function
		End If
	
	else  ''''' 폐기일 때, grid에 자산지출상세내역 입력시,삭제 
		
		if frm1.vspdData.MaxRows > 0 then 
			ggoSpread.Source = frm1.vspdData
			
			for intRow = 1 to frm1.vspdData.MaxRows 				
				frm1.vspdData.row = intRow
				lDelRows = ggoSpread.DeleteRow				
			next
			'frm1.vspdData.MaxRows = 0
			ggoSpread.Source = frm1.vspdData
			ggospread.ClearSpreadData		'Buffer Clear
			'''call InitSpreadSheet
		end if
	end if
	'-----------------------
	'Save function call area
	'-----------------------
	Call DbSave				                                                '☜: Save db data	
	FncSave = True
	
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()
    frm1.vspdData.ReDraw = False
    	
    if frm1.vspdData.MaxRows < 1 then Exit Function
    	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_RcptType	
	call subVspdSettingChange(frm1.vspdData.ActiveRow,frm1.vspdData.Text)

'	frm1.vspdData.Col = C_RcptType
'	frm1.vspdData.Text = ""

'	frm1.vspdData.Col = C_RcptTypeNm
'	frm1.vspdData.Text = ""

	MaxSpreadVal frm1.vspdData.ActiveRow
        
    frm1.vspdData.ReDraw = True
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 

	
	Dim iDx
	
	FncCancel = False

    if frm1.vspdData.MaxRows < 1 then Exit Function
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo
	
	Frm1.vspdData.Row = frm1.vspdData.ActiveRow
   	Frm1.vspdData.Col = C_RcptType
    iDx = Frm1.vspdData.Value
    Frm1.vspdData.Col = C_RcptTypeNm
    Frm1.vspdData.value = iDx
     
    Set gActiveElement = document.ActiveElement   
     
    FncCancel = True
		
	
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow(Byval pvRowCnt)

	Dim imRow, indx
	FncInsertRow = False
	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If

	With frm1
		if frm1.Rb_Sold.checked = True then		
	
			.vspdData.focus
			ggoSpread.Source = .vspdData
			'.vspdData.EditMode = True
			.vspdData.ReDraw = False
			ggoSpread.InsertRow ,imRow
			
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
    
        for indx = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
			'MaxSpreadVal .vspdData.ActiveRow
			call MaxSpreadVal(indx)
		next		
		.vspdData.ReDraw = True	
		
   		end if	
    End With
    Set gActiveElement = document.ActiveElement  
	If Err.number = 0 Then
	   FncInsertRow = True                                                          '☜: Processing is OK
	End If 
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    if frm1.vspdData.MaxRows < 1 then Exit Function

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
'		lgBlnFlgChgValue = True
    End With
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                              
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
    '완성이 되지 않았음 
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next
    '완성이 되지 않았음                                                    
End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)										
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
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


'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete() 
    Dim strVal
    
    DbDelete = False														'⊙: Processing is NG 
    
     Call LayerShowHide(1)  
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtChgNo=" & Trim(frm1.txtChgNo.value)			'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
	    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()												        '삭제 성공후 실행 로직 
	lgBlnFlgChgValue = False
	Call FncNew()
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
    
	DbQuery = False                                                         
	
	Call LayerShowHide(1)
	
	Dim strVal
	
	With frm1
	
        	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'☜: 
        	strVal = strVal     & "&txtChgNo=" & Trim(.txtChgNo.value)	'조회 조건 데이타 
        	strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey

			' 권한관리 추가 
			strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
			strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
			strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
			strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
			    	
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)										'비지니스 ASP 를 가동 
	
	DbQuery = True                                                          
    
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================================
Function DbQueryOk()													'조회 성공후 실행로직 
	Dim varData
	Dim iRow
	
	lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	
	Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field	
	Call SetToolBar("1111111100111111")									'버튼 툴바 제어 
	
	'Call InitData()			
	Call SetSpreadColor(-1,-1)
	''
	
	With frm1				
	
		.vspdData.Redraw = False
	
		For iRow = 1 To frm1.vspdData.MaxRows
	
			.vspdData.Col = C_RcptType		
			.vspdData.Row = iRow
			
			varData = frm1.vspdData.text
			
			Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type별 입력필수 필드 표시 
		Next
		
		.vspdData.Redraw = True			
	End With
	
	call txtDocCur_OnChangeASP()
	call txtCur_OnChange()
	Call txtVatAmt_Change()
	call txtVatType_OnChange()
	IF frm1.Rb_Duse.checked	= True Then
		Call radio2_onchange()
	END IF
	lgBlnFlgChgValue = False
    frm1.txtChgNo.focus 
	'SetGridFocus
	
End Function

Sub InitData()
	Dim intRow
	Dim intIndex 
	dim temp
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row  = intRow
			
			.Col	 = C_RcptType
			intIndex = .Value 			

			.Col     = C_RcptTypeNm
			.Value   = intindex					
		Next	
	End With	

End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	
	Dim IntRows 
	Dim IntCols 
	
	Dim lGrpcnt 
	Dim strVal
	Dim strDel
	
	DbSave = False                                                          
	
	'On Error Resume Next                                                   
	
	Call LayerShowHide(1)
	
	'Call SetSumItem
	
	strVal = ""
	strDel = ""
	
	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
	End With
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data 연결 규칙 
	' 0: Flag , 1: Row위치, 2~N: 각 데이타 
	
	lGrpCnt = 1

	With frm1.vspdData
	    
		For IntRows = 1 To .MaxRows
		
			.Row = IntRows
			.Col = 0

			If .Text = ggoSpread.DeleteFlag Then
				strDel = strDel & "D" & parent.gColSep & IntRows & parent.gColSep		'D=Delete

			ElseIf .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & parent.gColSep & IntRows & parent.gColSep		'U=Update

			Else				
				strVal = strVal & "C" & parent.gColSep & IntRows & parent.gColSep		'C=Create, Sheet가 2개 이므로 구별			
			End If
		
			Select Case .Text
				Case ggoSpread.DeleteFlag
					.Col = C_Seq
					strDel = strDel & Trim(.Text) & parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 
					
					lGrpcnt = lGrpcnt + 1             
								
				Case Else
					
					.Col = C_Seq
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_RcptType
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_Amt
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_LocAmt
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep					
					.Col = C_BankAcct
					strVal = strVal & Trim(.Text) & parent.gColSep					
					.Col = C_NoteNo
					strVal = strVal & Trim(.Text) & parent.gRowSep				    '마지막 데이타는 Row 분리기호를 넣는다 
					
					lGrpCnt = lGrpCnt + 1


			End Select
		Next

	End With

	frm1.txtMaxRows.value = lGrpCnt-1										'☜: Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value = strDel & strVal									'☜: Spread Sheet 내용을 저장 

	'권한관리추가 start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'권한관리추가 end
			
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'☜: 저장 비지니스 ASP 를 가동 

	DbSave = True                                                           
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

   	lgBlnFlgChgValue = false	

    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables                                                      'Initializes local global variables
'    Call InitSpreadSheet																			'⊙: Initializes local global variables
	'frm1.vspdData.MaxRows = 0 ' InitSpreadSheet 대신 
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    'Call InitComboBox
	Call DbQuery	

End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
Function Radio1_onChange	
    	
	ggoOper.SetReqAttr frm1.txtTotalAmt, "N"    '총입금액 
	ggoOper.SetReqAttr frm1.txtTotalLocAmt, "D"    '총입금액(자국)	
	ggoOper.SetReqAttr frm1.txtArAcctCd, "D"    '미수금계정 
	ggoOper.SetReqAttr frm1.txtVatRate, "D"    '부가세율 
	ggoOper.SetReqAttr frm1.txtVatAmt, "D"    '부가세금액(자국)
	ggoOper.SetReqAttr frm1.txtVatLocAmt, "D"    '부가세금액(자국)

	ggoOper.SetReqAttr frm1.txtReportAreaCd,		"D"    '신고사업장 
	ggoOper.SetReqAttr frm1.txtIssuedDt,	"D"    '발행일 
		
	ggoOper.SetReqAttr frm1.txtDueDt,		 "D"    '미수금만기일자			
	ggoOper.SetReqAttr frm1.txtBpCd,		 "N"    '거래처 
	ggoOper.SetReqAttr frm1.txtDocCur,	     "N"    '거래통화 
	frm1.txtDocCur.value = parent.gCurrency
	ggoOper.SetReqAttr frm1.txtVatType,		 "D"    '거래처 

    If lgIntFlgMode <> parent.OPMD_CMODE then                                   'Indicates that current mode is Create mode
		Call SetToolBar("1111111100111111")
		lgBlnFlgChgValue = True	
	Else
	    Call SetToolBar("1110110100100111")										' 처음 로드시 표준 에 따라 
	End if

End Function

Function Radio2_onChange
	Dim lDelRows,intRow
	Dim bMidChgVal
	
	ggoOper.SetReqAttr frm1.txtTotalAmt,	"Q"    '총입금액 
	ggoOper.SetReqAttr frm1.txtTotalLocAmt, "Q"    '총입금액(자국)	
	ggoOper.SetReqAttr frm1.txtArAcctCd, "Q"    '미수금계정 
	ggoOper.SetReqAttr frm1.txtApAmt, "Q"    '미수금금액 
	ggoOper.SetReqAttr frm1.txtApLocAmt, "Q"    '미수금금액(자국)
	ggoOper.SetReqAttr frm1.txtVatRate,		"Q"    '부가세율 
	ggoOper.SetReqAttr frm1.txtVatAmt,		"Q"    '부가세금액(자국)
	ggoOper.SetReqAttr frm1.txtVatLocAmt,	"Q"    '부가세금액(자국)

	ggoOper.SetReqAttr frm1.txtReportAreaCd, "Q"    '신고사업장 
	ggoOper.SetReqAttr frm1.txtIssuedDt,	 "Q"    '발행일 
		
	ggoOper.SetReqAttr frm1.txtDueDt,		 "Q"    '미수금만기일자 
	ggoOper.SetReqAttr frm1.txtBpCd,		 "Q"    '거래처 
	ggoOper.SetReqAttr frm1.txtDocCur,		 "Q"    '거래통화 
	ggoOper.SetReqAttr frm1.txtVatType,		 "Q"    '거래처 
	
	bMidChgVal = lgBlnFlgChgValue

	frm1.txtBpCd.value = ""
	frm1.txtBpNm.value = ""
	frm1.txtDocCur.value = ""
	
	frm1.txtApAmt.text		  = "0"
	frm1.txtApLocAmt.text	  = "0"
	
	if frm1.vspdData.MaxRows > 0 then 
		ggoSpread.Source = frm1.vspdData			
		frm1.vspdData.ReDraw = false	
		for intRow = 1 to frm1.vspdData.MaxRows 				
			frm1.vspdData.row = intRow
			lDelRows = ggoSpread.DeleteRow				
		next
		'frm1.vspdData.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear
		
		frm1.vspdData.ReDraw = True
	end if		
	
	lgBlnFlgChgValue = bMidChgVal
	
    If lgIntFlgMode <> parent.OPMD_CMODE then                              'Indicates that current mode is Create mode
		Call SetToolBar("1111100000011111")									'버튼 툴바 제어	
		lgBlnFlgChgValue = True	
	Else
	    Call SetToolBar("1110100000000111")	
	End if
	

End Function

function txtDeptCd_onblur()
	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
end function
Sub txtDeptCd_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtRegDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
	
		'----------------------------------------------------------------------------------------

End Sub
Sub txtChgDt_onBlur()
    
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2

	lgBlnFlgChgValue = True
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtChgDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
				strSelect	=			 " Distinct org_change_id "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hORGCHANGEID.value) Then
					'IntRetCD = DisplayMsgBox("124600","X","X","X") 
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hORGCHANGEID.value = ""
					.txtDeptCd.focus
			End if
		End If
	End With
'----------------------------------------------------------------------------------------

End Sub


function txtBpCd_onblur()
	if Trim(frm1.txtBpCd.value) = "" then 		
		frm1.txtBpNm.value = ""		
	end if	
End function

Function txtDueDt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtChgDt_Change()
	Dim StrDeptCd
	
	StrDeptCd = Trim(frm1.txtDeptCd.value)
	
	if lgBlnFlgChgValue = true and StrDeptCd <> "" then Call txtDeptCd_onchange()
	lgBlnFlgChgValue = True
End Function

Function txtIssuedDt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtXchRate_Change()
	lgBlnFlgChgValue = True
End Function

Function txtChgAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtChgLocAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtTotalAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtTotalLocAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtApAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtApLocAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtVatRate_Change()
	lgBlnFlgChgValue = True
End Function

Function txtVatLocAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtChgQty_Change()
	lgBlnFlgChgValue = True
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1

'		ggoOper.FormatFieldByObjectOfCur .txtAcqAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtTotalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtApAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtDeprAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec

	End With

End Sub
'===================================== CurFormatNumericOCXRef()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCXRef()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtAcqAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With

End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_Amt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub
'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!--
'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenMasterRef()">자산마스터참조</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>					
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
									<TD CLASS="TD5" NOWRAP>자산변동번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtChgNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="자산변동번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenChgNoInfo"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=20 valign=top>
						<TABLE <%=LR_SPACE_TYPE_50%>>	
									<TR>
										<TD CLASS="TD5" NOWRAP>자산번호</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNo" SIZE=20 MAXLENGTH=18 TAG="24XXXU" ALT="자산번호"> <INPUT TYPE="Text" NAME="txtAsstNm" SIZE=25 MAXLENGTH=40 tag="24X" ALT="자산명"></TD>
										<TD CLASS="TD5" NOWRAP>취득일자</TD>
										<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtRegDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="취득일자" tag="24X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>관리부서</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctDeptNm" SIZE=27 MAXLENGTH=40 tag="24XXXU" ALT="관리부서명"></TD>
										<TD CLASS="TD5" NOWRAP>취득-재고수량</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle0 name="txtAcqQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="취득-재고수량" tag="24X3P"> </OBJECT>');</SCRIPT>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name="txtInvQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="취득-재고수량" tag="24X3P"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>거래통화|환율</TD>
										<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCur" ALT="거래통화" TYPE="Text" MAXLENGTH=3 SIZE=10 tag="24XXXU" >&nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtXchRt" align="top" CLASS=FPDS90 title="FPDOUBLESINGLE" ALT="환율" tag="24X5Z"></OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD5 NOWRAP>취득금액|자국</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtAcqAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="취득금액" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 name=txtAcqLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="취득금액(자국)" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>
									</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>매각/폐기일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtChgDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="매각/폐기일자" tag="22X1" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>회계부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="회계부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDept()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 tag="24" ALT="회계부서명"></TD>
							</TR>
							<TR>
					        <TD CLASS="TD5" NOWRAP>매각/폐기구분</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_Sold Checked tag = 2 value="03" onclick=radio1_onchange()><LABEL FOR=Rb_Sold>매각</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Duse tag = 2 value="04" onclick=radio2_onchange()><LABEL FOR=Rb_Duse>폐기</LABEL></TD>
								<TD CLASS=TD5 NOWRAP>매각/폐기수량</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle10 name=txtChgQty style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="매각/폐기 수량" tag="22X3P"></OBJECT>');</SCRIPT>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>거래통화|환율</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" SIZE=10 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurrency()">&nbsp;
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtXchRate align="top" CLASS=FPDS90 title=FPDOUBLESINGLE ALT="환율" tag="22X5Z" id=fpDoubleSingle5></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS="TD5" NOWRAP>거래처</TD>
	                            <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24" ALT="거래처명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>총입금액</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="총입금액" tag="22X2" id=fpDoubleSingle6></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>총입금액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="총입금액(자국)" tag="21X2" id=fpDoubleSingle7></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>미수금계정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtArAcctCd" SIZE=10 MAXLENGTH=20 tag="21XXXU" ALT="미수금계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnArAcctCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenArAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtArAcctNm" SIZE=22 tag="24"  alt = "미수금계정명"></TD>
								<TD CLASS=TD5 NOWRAP>미수금액|자국</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtApAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="미수금액" tag="24X2" id=fpDoubleSingle8></OBJECT>');</SCRIPT>&nbsp;
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtApLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="미수금액(자국)" tag="24X2" id=fpDoubleSingle9></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>미수금 만기일자</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtDueDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21X1" ALT="미수금 만기일자"> </OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS="TD5" NOWRAP>미수금 번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtApNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="미수금 번호"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>감가상각누계금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle11 name=txtDeprAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="감가상각누계금액" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
								</TD>
								<TD CLASS=TD5 NOWRAP>감가상각누계금액(자국)</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle12 name=txtDeprLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="감가상각누계금액(자국)" tag="24X2"> </OBJECT>');</SCRIPT> &nbsp;
	                            </TD>
							</TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP>부가세유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="부가세유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType()">&nbsp;<INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="24" ALT="부가세유형"></TD>
								<TD CLASS="TD5" NOWRAP>부가세율</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtVatRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="부가세율" tag="21X5Z" id=fpDoubleSingle10></OBJECT>');</SCRIPT>&nbsp;%</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부가세금액|자국</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtVatAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="부가세금액" tag="21X2" id=fpDoubleSingle11> </OBJECT>');</SCRIPT>&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtVatLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="부가세금액(자국)" tag="21X2" id=fpDoubleSingle12> </OBJECT>');</SCRIPT> &nbsp;
	                            </TD>
	                            <TD CLASS="TD5" NOWRAP>계산서발행일</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 name=txtIssuedDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21" ALT="전표생성일자"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>신고사업장</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtReportAreaCd" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="신고사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReportAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReportAreaCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtReportAreaNm" SIZE=20 tag="24" ALT="신고사업장명"></TD>
								<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="결의전표번호"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>적요</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtChgDesc" SIZE=35 MAXLENGTH=30 tag="2X" ALT="적요"></TD>
								<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="회계전표번호"></TD>
							</TR>
							<TR>
								<TD WIDTH="80%" HEIGHT=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData HEIGHT="100%" tag="2" width="100%" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>

			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"         tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"   tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"  tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows"	  tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode"	  tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hORGCHANGEID"   tag="24"TABINDEX = "-1" >
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
