<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 보조부조회 
'*  3. Program ID           : A5111MA1
'*  4. Program Name         : 보조부조회 
'*  5. Program Desc         : 
'*  6. Component List       :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2004/01/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Chang Jin
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'========================================================================================================

Const BIZ_PGM_ID 	   = "a5111mb1.asp"                              '☆: Biz Logic ASP Name
Const BIZ_PGM_POPUP_ID = "a5111mb2.asp"	

'========================================================================================================
Const C_MaxKey          = 6					                          '☆: SpreadSheet의 키의 갯수 

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
Dim lgIsOpenPop
Dim lgMaxFieldCount
Dim lgCookValue 
Dim lgSaveRow 

Dim BaseDate
Dim LastDate, FirstDate
Dim FromDateOfDB, ToDateOfDB

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

   BaseDate     = "<%=GetSvrDate%>"                                                                  'Get DB Server Date

   LastDate     = UNIGetLastDay (BaseDate,Parent.gServerDateFormat)                                  'Last  day of this month
   FirstDate    = UNIGetFirstDay(BaseDate,Parent.gServerDateFormat)                                  'First day of this month

   FromDateOfDB = UNIDateAdd("yyyy", -20, BaseDate,Parent.gServerDateFormat)
   ToDateOfDB   = UNIDateAdd("yyyy", 410, BaseDate,Parent.gServerDateFormat)
 
   FromDateOfDB  = UniConvDateAToB(FromDateOfDB ,Parent.gServerDateFormat,Parent.gDateFormat)
   ToDateOfDB    = UniConvDateAToB(ToDateOfDB   ,Parent.gServerDateFormat,Parent.gDateFormat)



'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
End Sub

'========================================================================================================
Sub SetDefaultVal()
	Dim StartDate, EndDate
	Dim strYear, strMOnth, strDay

	Call ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate= UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")		'☆: 초기화면에 뿌려지는 시작 날짜 
	EndDate= UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)		'☆: 초기화면에 뿌려지는 마지막 날짜 

    frm1.txtFromGlDt.text	= StartDate 
	frm1.txtToglDt.Text		= EndDate 

	Call ElementVisible(frm1.txtSubLedger1, 0)
	Call ElementVisible(frm1.btnSubLedger1, 0)
	Call ElementVisible(frm1.txtSubLedger3, 0)
	Call ElementVisible(frm1.txtSubLedger2, 0)
	Call ElementVisible(frm1.btnSubLedger2, 0)
	Call ElementVisible(frm1.txtSubLedger4, 0)
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "*", "COOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "*", "COOKIE", "QA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)
	Dim strTemp, arrVal

	Const CookieSplit = 4877

	If Kubun = 0 Then
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" Then Exit Function

       arrVal = Split(strTemp, Parent.gRowSep)

       Frm1.txtSchoolCd.Value = ReadCookie ("SchoolCd")
       Frm1.txtGrade.Value   = arrVal(0)

       Call MainQuery()

       WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 1 Then                                         ' If you want to call
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End If
End Function

'========================================================================================
Sub InitSpreadSheet()
 	Call SetZAdoSpreadSheet("A5111MA1Q01","S","A","V20021220",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,	-1
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")	
    
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


    frm1.txtFromGlDt.focus
    Call CookiePage(0)
    
    frm1.txtTDrAmt.allownull = False
    frm1.txtTCrAmt.allownull = False
    frm1.txtTSumAmt.allownull = False

    frm1.txtNDrAmt.allownull = False
    frm1.txtNCrAmt.allownull = False
    frm1.txtNSumAmt.allownull = False

    frm1.txtSDrAmt.allownull = False
    frm1.txtSCrAmt.allownull = False
    frm1.txtSSumAmt.allownull = False
End Sub

'========================================================================================
Function FncQuery() 
    FncQuery = False
    Err.Clear

    '-----------------------
    'Erase contents area
    '-----------------------
	If Trim(Frm1.txtSubLedger1.value) = "" Then
		Frm1.txtSubLedger3.value = ""
	End If

	If Trim(Frm1.txtSubLedger2.value) = "" Then
		Frm1.txtSubLedger4.value = ""
	End If
    
    Call InitVariables
      
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
		Exit Function
    End If

    If UniConvDateToYYYYMMDD(frm1.txtFromGlDt.text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToGlDt.Text, Parent.gDateFormat, "") Then
		Call SetToolBar("1100000000001111")	
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If

    Call ggoOper.ClearField(Document, "2")	
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True
End Function

'========================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear
	Call Parent.FncPrint()
    FncPrint = True
End Function

'========================================================================================
Function FncExcel() 
    FncExcel = False
    Err.Clear
	Call Parent.FncExport(Parent.C_MULTI)
    FncExcel = True
End Function

'========================================================================================
Function FncFind() 
    FncFind = False
    Err.Clear
	Call Parent.FncFind(Parent.C_MULTI, True)
    FncFind = True
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
    FncExit = False
    Err.Clear
    FncExit = True
End Function

'========================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear
    DbQuery = False

	Call LayerShowHide(1)

    With frm1
        strVal = BIZ_PGM_ID
        'If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
		
			strVal = strVal & "?txtFromGlDt=" & UniConvDateToYYYYMMDD(frm1.txtFromGlDt.Text,Parent.gDateFormat,"")
			strVal = strVal & "&txtToGlDt="   & UniConvDateToYYYYMMDD(frm1.txtToGlDt.Text,Parent.gDateFormat,"")
			strVal = strVal & "&txtAcctCd=" & .txtAcctCd.value
			strVal = strVal & "&txtBizAreaCd=" & .txtBizAreaCd.value				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtBizAreaCd1=" & .txtBizAreaCd1.value
			strVal = strVal & "&txtSubLedger1=" & .txtSubLedger1.value
			strVal = strVal & "&txtSubLedger2=" & .txtSubLedger2.value
			strVal = strVal & "&txtMajorCd1=" & .hMajorCd1.value
			strVal = strVal & "&txtMajorCd2=" & .hMajorCd2.value
	    'Else
		'	strVal = strVal & "?txtFromGlDt=" & Trim(.hFromGlDt.value)
		'	strVal = strVal & "&txtToGlDt=" & Trim(.hToGlDt.value)
        '   strVal = strVal & "&txtAcctCd=" & Trim(.hAcctCd.value)
        '   strVal = strVal & "&txtBizAreaCd=" & Trim(.hBizAreaCd.value)			'☆: 조회 조건 데이타 
        '   strVal = strVal & "&txtBizAreaCd1=" & Trim(.hBizAreaCd1.value)
		'	strVal = strVal & "&txtSubLedger1=" & Trim(.hSubLedger1.value)
		'	strVal = strVal & "&txtSubLedger2=" & Trim(.hSubLedger2.value)
		'	strVal = strVal & "&txtMajorCd1=" & Trim(.hMajorCd1.value)
		'	strVal = strVal & "&txtMajorCd2=" & Trim(.hMajorCd2.value)
        'End If
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

		Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()												
	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1 

    Call SetToolBar("1100000000011111")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
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
	
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	
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
	
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

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
		Call SetReturnVal(arrRet,5)
	End If
End Function

'----------------------------------------  OpenAcctCd()  -------------------------------------------------
'	Name : OpenAcctCd()
'	Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "계정 팝업"	
	arrParam(1) = " A_ACCT A, A_CTRL_ITEM B, A_CTRL_ITEM C, A_ACCT_GP D "
	arrParam(2) = Trim(frm1.txtAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.SUBLEDGER_1 *= B.CTRL_CD AND A.SUBLEDGER_2 *= C.CTRL_CD AND A.GP_CD = D.GP_CD AND (A.SUBLEDGER_1 IS NOT NULL AND A.SUBLEDGER_1 <> '') "
	arrParam(5) = "계정코드"

	arrField(0) = "A.ACCT_CD"								' Field명(0)
	arrField(1) = "A.ACCT_NM"								' Field명(1)
	arrField(2) = "D.GP_NM"									' Field명(2)
	arrField(3) = "B.CTRL_CD"									' Field명(3)
	arrField(4) = "B.CTRL_NM"
	arrField(5) = "C.CTRL_CD"
	arrField(6) = "C.CTRL_NM"


	arrHeader(0) = "계정코드"
	arrHeader(1) = "계정명"
	arrHeader(2) = "그룹명"							' Header명(2)
	arrHeader(3) = "관리항목1"								' Header명(3)
	arrHeader(4) = "관리항목명1"
	arrHeader(5) = "관리항목2"
	arrHeader(6) = "관리항목명2"

   arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,2)
	End If	
End Function

'-------------------------------------  OpenSubLedger()  -------------------------------------------------
'	Name : OpenSubLedger()
'	Description : Open SubLedger PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSubLedger(byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function

	if Trim(frm1.txtAcctCd.value) = "" then
		Call DisplayMsgBox("700102", "X", "X", "X")                                '☆: 밑에 메세지를 ID로 처리해야 함 
		'계정과목을 입력하세요.
	end if

	lgIsOpenPop = True

	arrParam(0) = "관리항목 팝업"	' 팝업 명칭 
	arrParam(3) = ""					' Name Condition

	If iWhere = 1 And Trim(frm1.hBTblId.value) = "" Then
		lgIsOpenPop = False
	ElseIf iWhere = 2 And Trim(frm1.hCTblId.value) = "" Then
		lgIsOpenPop = False
	Else
		If iWhere = 1 Then
			arrParam(1) = Trim(frm1.hBTblId.Value)			' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtSubLedger1.Value)	' Code Condition
			arrParam(4) = Trim(frm1.hBColmId.Value) & " IN (SELECT CTRL_VAL1 FROM A_SUBLEDGER_SUM WHERE ACCT_CD = " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ")"

			If UCase(Trim(frm1.hBTblId.Value)) = "B_MINOR" Then
				arrParam(4) = arrParam(4) & " AND MAJOR_CD = " & FilterVar(frm1.hMajorCd1.value, "''", "S")  
			End If

			arrParam(5) =  lblTitle1.innerHTML 

			arrField(0) = Trim(frm1.hBColmId.Value)			' Field명(0)
			arrField(1) = Trim(frm1.hBColmIdNm.Value)		' Field명(1)
	    Else
			arrParam(1) = Trim(frm1.hCTblId.Value)
			arrParam(2) = Trim(frm1.txtSubLedger2.Value)
			arrParam(4) = Trim(frm1.hCColmId.Value) & " IN (SELECT CTRL_VAL2 FROM A_SUBLEDGER_SUM WHERE ACCT_CD = " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ")"

			If UCase(Trim(frm1.hCTblId.Value)) = "B_MINOR" Then
				arrParam(4) = arrParam(4) & " AND MAJOR_CD = " & FilterVar(frm1.hMajorCd2.value, "''", "S")  
			End If

			arrParam(5) =  lblTitle2.innerHTML 

			arrField(0) = Trim(frm1.hCColmId.Value)
			arrField(1) = Trim(frm1.hCColmIdNm.Value)
	    End If

		arrHeader(0) = Trim(frm1.hBColmId.Value)			' Header명(0)
		arrHeader(1) = Trim(frm1.hBColmIdNm.Value)			' Header명(1)

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False

		If arrRet(0) = "" Then
			If iWhere = 1 Then
				frm1.txtSubLedger1.focus
			Else
				frm1.txtSubLedger2.focus
			End If
			Exit Function
		Else
			If iWhere = 1 Then
				Call SetReturnVal(arrRet,3)
			Else
				Call SetReturnVal(arrRet,4)
			End If
		End If	
	End If
End Function

'-------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(ByVal arrRet,ByVal field_fg) 
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim ii, jj
	Dim arrVal1, arrVal2

	With frm1	
		Select case field_fg
			case 1
'				.txtBizAreaCd.focus
				.txtBizAreaCd.Value		= arrRet(0)
				.txtBizAreaNm.Value		= arrRet(1)
			case 2
				.txtSubLedger1.value    = ""
				.txtSubLedger2.value    = ""
				.txtSubLedger3.value    = ""
				.txtSubLedger4.value    = ""

'				.txtAcctCd.focus
				.txtAcctCd.Value		= arrRet(0)
				.txtAcctNm.Value		= arrRet(1)

				.hBCtrlCd.Value			= arrRet(3)
				lblTitle1.innerHTML		= arrRet(4)
				.hCCtrlCd.Value			= arrRet(5)
				lblTitle2.innerHTML		= arrRet(6)
'				.hMajorCd1.Value		= arrRet(8)
'				.hMajorCd2.Value		= arrRet(9)

	'----------------------------------------------------------------------------------------
	' MajorCD1, MajorCd2 가져오기 
				strSelect	=			 " b.major_cd, c.major_cd   "    		
				strFrom		=			 " A_ACCT A, A_CTRL_ITEM B, A_CTRL_ITEM C, A_ACCT_GP D "		
				strWhere	=			 " A.SUBLEDGER_1 *= B.CTRL_CD AND A.SUBLEDGER_2 *= C.CTRL_CD AND A.GP_CD = D.GP_CD"	
				strWhere	= strWhere & " AND (A.SUBLEDGER_1 IS NOT NULL AND A.SUBLEDGER_1 <> '') And A.ACCT_CD= "
				strWhere	= strWhere & FilterVar(LTrim(RTrim(.txtAcctCd.Value)), "''", "S")

					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
						jj = Ubound(arrVal1,1)

						For ii = 0 To jj - 1

							arrVal2			= Split(arrVal1(ii), chr(11))
							.hMajorCd1.Value			= Trim(arrVal2(1))
							.hMajorCd2.Value			= Trim(arrVal2(2))

						Next
					end if 
	'----------------------------------------------------------------------------------------
				Call DbPopUpQuery()
			case 3
'				.txtSubLedger1.focus
				.txtSubLedger1.value	= arrRet(0)
				.txtSubLedger3.value	= arrRet(1)
			case 4	'OpenSubledger2
'				.txtSubLedger2.focus
				.txtSubLedger2.value	= arrRet(0)
				.txtSubLedger4.value	= arrRet(1)
			case 5
				.txtBizAreaCd1.Value	= arrRet(0)
				.txtBizAreaNm1.Value	= arrRet(1)
		End select
	End With
End Function

'========================================================================================================
'	Name : DbPopUpQuery()
'	Description : 
'========================================================================================================
Function DbPopUpQuery()
    Dim strVal

'    Err.Clear
    DbPopUpQuery = False                                                     '⊙: Processing is NG

	Call LayerShowHide(1)

'    vspdData.MaxRows = 0

    strVal = BIZ_PGM_POPUP_ID & "?"											'☜: 
	strVal = strVal & "hBCtrlCd=" & Trim(frm1.hBCtrlCd.value)				'☆: 조회 조건 데이타 
	strVal = strVal & "&hCCtrlCd=" & Trim(frm1.hCCtrlCd.value)				'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbPopUpQuery = True
End Function

'========================================================================================================
'	Name : DbPopUpQueryOk()
'	Description : 
'========================================================================================================
Function DbPopUpQueryOk()
    DbPopUpQueryOk = False

	If Trim(frm1.hBCtrlCd.value) = "" Then
		Call ElementVisible(frm1.txtSubLedger1, 0)
		Call ElementVisible(frm1.btnSubLedger1, 0)
		Call ElementVisible(frm1.txtSubLedger3, 0)
	Else
		frm1.txtSubLedger1.disabled = False
		Call ElementVisible(frm1.txtSubLedger1, 1)
		Call ElementVisible(frm1.btnSubLedger1, 1)
		Call ElementVisible(frm1.txtSubLedger3, 1)
	End If

	If Trim(frm1.hCCtrlCd.value) = "" Then
		Call ElementVisible(frm1.txtSubLedger2, 0)
		Call ElementVisible(frm1.btnSubLedger2, 0)
		Call ElementVisible(frm1.txtSubLedger4, 0)
	Else
		Call ElementVisible(frm1.txtSubLedger2, 1)
		Call ElementVisible(frm1.btnSubLedger2, 1)
		Call ElementVisible(frm1.txtSubLedger4, 1)
	End If

    DbPopUpQueryOk = True
End Function

'-------------------------------------  OpenPopupGL()  --------------------------------------------------
'	Name : OpenPopupGL()
'	Description :
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupGL() 
	Dim arrRet
	Dim arrParam(1)	
	Dim arrField
	Dim intFieldCount
	Dim i
	Dim j

	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData									'z_ado_field_inf의 내용이 바뀌면..이곳을 반드시 확인해야한다.
		If .maxrows > 0 Then
			arrField = Split(GetSQLSelectListDataType("A"),",")
			intFieldCount = UBound(arrField,1)
			For i = 0 To  intFieldCount -1
				If Trim(arrField(i)) = "C.GL_NO" Then
					Exit For
				End if
			Next
		
			.Row = .ActiveRow
			.Col = i + 2

			arrParam(0) = Trim(.Text)	'결의전표번호 
			arrParam(1) = ""			'Reference번호 
		End if	
	End With

	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
End Function

'========================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet

	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "2")
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()
   End If
End Function


'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub

'========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function

'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"	

	Set gActiveSpdSheet = frm1.vspdData

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If
    End If

    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
   	End If

	lgCookValue = ""
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row) 
End Sub

'========================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
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

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromGlDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToGlDt.Focus
    End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub txtFromGlDt_Keypress(Key)
    If Key = 13 Then
'		frm1.txtToGlDt.focus
        MainQuery()
    End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub txtToGlDt_Keypress(Key)
    If Key = 13 Then
'		frm1.txtFromGlDt.focus
        MainQuery()
    End If
End Sub

'==========================================================================================
'   Event Name : txtAcctCd_OnBlur
'   Event Desc : 
'==========================================================================================
Sub txtAcctCd_OnBlur()
	Dim ArrRet
	Dim ArrParam(7)

	IF Trim(frm1.txtAcctCd.value) = "" Then
		Exit sub
	End If

    Call CommonQueryRs( "A.ACCT_CD,A.ACCT_NM,D.GP_NM,B.CTRL_CD,B.CTRL_NM,C.CTRL_CD,C.CTRL_NM" , _ 
				"A_ACCT A, A_CTRL_ITEM B, A_CTRL_ITEM C, A_ACCT_GP D " , _
				 "A.ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & " AND A.SUBLEDGER_1 *= B.CTRL_CD AND A.SUBLEDGER_2 *= C.CTRL_CD AND A.GP_CD = D.GP_CD AND (A.SUBLEDGER_1 IS NOT NULL AND A.SUBLEDGER_1 <> '') " , _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If lgF0 = "" Then
		Exit Sub
	End If

	ArrRet 	= Split(lgF0,Chr(11))
	ArrParam(0) = ArrRet(0)
	ArrRet 	= Split(lgF1,Chr(11))
	ArrParam(1) = ArrRet(0)
	ArrRet 	= Split(lgF2,Chr(11))
	ArrParam(2) = ArrRet(0)
	ArrRet 	= Split(lgF3,Chr(11))
	ArrParam(3) = ArrRet(0)
	ArrRet 	= Split(lgF4,Chr(11))
	ArrParam(4) = ArrRet(0)
	ArrRet 	= Split(lgF5,Chr(11))
	ArrParam(5) = ArrRet(0)

	Call SetReturnVal(arrParam,2)
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="회계일자" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
												           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="회계일자" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd()">&nbsp;
									                       <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd1()">&nbsp;
									                       <INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12X5Z" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenAcctCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle1">&nbsp;</SPAN></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSubLedger1" SIZE=16 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="관리항목1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSubLedger1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSubLedger(1)">&nbsp;
									                       <INPUT TYPE=TEXT NAME="txtSubLedger3" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle2">&nbsp;</SPAN></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSubLedger2" SIZE=16 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="관리항목2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSubLedger2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSubLedger(2)">&nbsp;
									                       <INPUT TYPE=TEXT NAME="txtSubLedger4" SIZE=20 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
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
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="이월금액(차변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="이월금액(대변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="이월금액(자국)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발생금액</TD>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="발생금액(차변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="발생금액(대변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="발생금액(자국)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>누계금액</TD>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="누계금액(차변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="누계금액(대변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="누계금액(자국)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
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
	<TR>		
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1" ></IFRAME>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode"     tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hBCtrlCd"    tag="14" TABINDEX="-1" ><INPUT TYPE=HIDDEN NAME="hBTblId" tag="14" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hBColmId" tag="14" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hBColmIdNm" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCCtrlCd"    tag="14" TABINDEX="-1" ><INPUT TYPE=HIDDEN NAME="hCTblId" tag="14" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hCColmId" tag="14" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hCColmIdNm" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBizAreaCd"  tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hBizAreaCd1"  tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hFromGlDt"   tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hToGlDt"     tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hAcctCd"     tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hSubLedger1" tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hSubLedger2" tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hMajorCd1" tag="14" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hMajorCd2" tag="14" TABINDEX="-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

