<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5131MA1
'*  4. Program Name         : 기초분개장조회 
'*  5. Program Desc         : Ado query Sample with DBAgent(Sort)
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/23
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Jung Sung Ki
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
	
'========================================================================================

Const BIZ_PGM_ID 		= "a5131mb1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================
Const C_MaxKey          = 1					                          '☆: SpreadSheet의 키의 갯수 

'========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================
Dim lgIsOpenPop                                          

Dim lgSelectList                                         
Dim lgSelectListDT                                       

Dim lgCookValue 


Dim lgSaveRow 

Dim IsOpenPop 

Dim BaseDate
Dim LastDate, FirstDate
Dim FromDateOfDB, ToDateOfDB

                                                 
   BaseDate     = "<%=GetSvrDate%>"                                                                  'Get DB Server Date
'  BaseDate     = Date(You must not code like this!!!!)                                       'Get AP Server Date

   LastDate     = UNIGetLastDay (BaseDate,Parent.gServerDateFormat)                                  'Last  day of this month
   FirstDate    = UNIGetFirstDay(BaseDate,Parent.gServerDateFormat)                                  'First day of this month

   FromDateOfDB = UNIDateAdd("yyyy", -410, BaseDate,Parent.gServerDateFormat)
   ToDateOfDB   = UNIDateAdd("yyyy",  410, BaseDate,Parent.gServerDateFormat)
 
   FromDateOfDB  = UniConvDateAToB(FromDateOfDB ,Parent.gServerDateFormat,Parent.gDateFormat)               '
   ToDateOfDB    = UniConvDateAToB(ToDateOfDB   ,Parent.gServerDateFormat,Parent.gDateFormat)               '



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

	Dim StartDate
	Dim strYear, strMonth, strDay

	Call	ExtractDateFrom(BaseDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate= UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")		'☆: 초기화면에 뿌려지는 시작 날짜 
'	EndDate= UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)		'☆: 초기화면에 뿌려지는 마지막 날짜 

	frm1.txtAccountYear.text	=	StartDate 
	Call ggoOper.FormatDate(frm1.txtAccountYear, Parent.gDateFormat, 3)

	frm1.hOrgChangeId.value = parent.gChangeOrgId
	frm1.txtBizArea.value	= Parent.gBizArea
	'frm1.txtAmtFr.Text	= ""
	'frm1.txtAmtTo.Text	= ""

	frm1.txtAccountYear.focus
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "QA") %>
End Sub


'========================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		Dim strTemp, arrVal

	Const CookieSplit = 4877						

	If Kubun = 0 Then                                              ' Called Area
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, Parent.gRowSep)


       WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF

	
End Function

'========================================================================================
Sub InitComboBox()
	
	'Err.clear
	'Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd='A1001' and (minor_cd ='BR' or minor_cd ='TR' or SUBSTRING(minor_cd,2,1) = 'T' ) order by minor_nm", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	'Call SetCombo2(frm1.txtGlInputType ,lgF0  ,lgF1  ,Chr(11))

End Sub
 


'========================================================================================
Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("A5120QA1","S","A","V20020101",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	    Call SetSpreadLock 
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")

	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")
    Call InitComboBox()
    Call CookiePage(0)
End Sub
'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub


'========================================================================================
Sub txtAccountYear_DblClick(Button)
    If Button = 1 Then  
        frm1.txtAccountYear.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtAccountYear.Focus
    End If
End Sub

'========================================================================================
Sub txtAccountYeart_Change() 
    lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub txtAccountYear_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub





'========================================================================================
Function FncQuery() 
	Dim IntRetCD
    FncQuery = False

    Err.Clear

    Call ggoOper.ClearField(Document, "2")

    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
 '#   If CompareDateByFormat(frm1.txtFromGlDt.text,frm1.txtToGlDt.text,frm1.txtFromGlDt.Alt,frm1.txtToGlDt.Alt, _
 '                       "970025",frm1.txtFromGlDt.UserDefinedFormat,Parent.gComDateType,True) = False Then			
'		Exit Function
 '   End If


    If frm1.txtBizArea.value = "" Then
		frm1.txtBizAreaNm.value = ""
    End If

    If frm1.txtdeptcd.value = "" Then
		frm1.txtdeptnm.value = ""
    End If
	IF NOT CheckOrgChangeId Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
		Exit Function
	End if

    IF DbQuery	 = False Then
       Exit Function
    End IF
       
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
	Dim IntRetCD

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
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode = Parent.OPMD_CMODE Then   ' This means that it is first search
        
			strVal = strVal & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtAccountYear=" & Trim(.txtAccountYear.text)
			strVal = strVal & "&txtBizArea=" & UCase(Trim(.txtBizArea.value))
			strVal = strVal & "&txtdeptcd=" & UCase(Trim(.txtdeptcd.value))				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtCOST_CENTER_CD=" & Trim(.txtCOST_CENTER_CD.value)
			strVal = strVal & "&txtRefNo=" & .txtRefNo.value
			strVal = strVal & "&txtGlInputType=" & .txtGlInputType.value
			strVal = strVal & "&txtDesc=" & Trim(.txtDesc.Value)
        Else
            strVal = strVal & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtAccountYear=" & Trim(.htxtAccountYear.value)
			strVal = strVal & "&txtBizArea=" & UCase(Trim(.htxtBizArea.value))
			strVal = strVal & "&txtdeptcd=" & UCase(Trim(.htxtdeptcd.value))				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtCOST_CENTER_CD=" & Trim(.htxtCOST_CENTER_CD.value)
			strVal = strVal & "&txtRefNo=" & .htxtRefNo.value
			strVal = strVal & "&txtGlInputType=" & .htxtGlInputType.value
			strVal = strVal & "&txtDesc=" & Trim(.htxtDesc.Value)
       End If
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&hChangeOrgId="       & .horgchangeid.value
        strVal = strVal & "&lgPageNo="       & lgPageNo
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()

	lgBlnFlgChgValue = False
    lgIntFlgMode     = Parent.OPMD_UMODE
    lgSaveRow        = 1
	Call SetToolbar("1100000000011111")	
	frm1.vspdData.focus
	CALL vspdData_Click(1, 1)
End Function


'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
   
		Case 1
			arrParam(0) = "사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""						' Where Condition
			arrParam(5) = "사업장코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"						' Field명(0)
			arrField(1) = "BIZ_AREA_NM"						' Field명(1)

			arrHeader(0) = "사업장코드"			' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

		Case 2
			arrParam(0) = "코스트센타 팝업"					' 팝업 명칭 
			arrParam(1) = "B_COST_CENTER"							' TABLE 명칭 
			arrParam(2) = strCode			       				    ' Code Condition
			arrParam(3) = ""										' Name Cindition
			arrParam(4) = ""										' Where Condition
			arrParam(5) = "코스트센타"

		    arrField(0) = "COST_CD"									' Field명(0)
			arrField(1) = "COST_NM"									' Field명(1)

			arrHeader(0) = "코스트센타코드"					' Header명(0)
			arrHeader(1) = "코스트센타명"						' Header명(1)	
		Case 3
			arrParam(0) = "전표입력경로팝업"
			arrParam(1) = "B_MINOR"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD = " & FilterVar("A1001", "''", "S") & " "
			arrParam(5) = "전표입력경로코드"

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"

			arrHeader(0) = "전표입력경로코드"
			arrHeader(1) = "전표입력경로명"
End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1		' Biz area
				frm1.txtBizArea.focus
			Case 2		' Biz area
				frm1.txtCOST_CENTER_CD.focus
			Case 3		' Biz area
				frm1.txtGlInputType.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If

End Function

'========================================================================================
'	Name : SetAcct()
'	Description : Account Popup에서 Return되는 값 setting
'========================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0	     'DEPT
			Case 1		' Biz area
				.txtBizArea.focus
				.txtBizArea.value		= UCase(Trim(arrRet(0)))
				.txtBizAreaNm.value		= arrRet(1)
			Case 2		' cost center
				.txtCOST_CENTER_CD.focus
				.txtCOST_CENTER_CD.value		= UCase(Trim(arrRet(0)))
				.txtCOST_CENTER_NM.value		= arrRet(1)
			Case 3		' Gl input type
				.txtGlInputType.focus
				.txtGlInputType.value		= UCase(Trim(arrRet(0)))
				.txtGLInputTypeNm.value		= arrRet(1)
		End Select
	End With
End Function

'========================================================================================
Function OpenDept()
	Dim arrRet
	Dim arrParam(4)
	Dim Temp
	Dim strYear, strMonth, strDay
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Call parent.ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
	arrParam(0) = frm1.txtDeptCd.value		            '  Code Condition
	arrParam(0) = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtAccountYear.text, "01", "01")
	arrParam(1) = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtAccountYear.text, "12", "31")
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "T"									' 결의일자 상태 Condition  


	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtdeptcd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
			
End Function

'========================================================================================
Function SetDept(ByVal arrRet)
		
	With frm1
		.txtDeptCd.focus
		.txtDeptCd.value = Trim(arrRet(0))
		.txtDeptNm.value = arrRet(1)
		.hOrgChangeId.value=arrRet(2)
	End With
End Function 


'========================================================================================
'	Name : OpenPopupGL()
'	Description :
'========================================================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)
	Dim IntRetCD
	Dim iCalledAspName


	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	arrParam(0) = Trim(GetKeyPosVal("A", 1))	'전표번호 
	arrParam(1) = ""
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	frm1.vspdData.focus
	IsOpenPop = False

End Function

'========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'========================================================================================
Sub txtDeptCD_OnChange()

    Dim strSelect, strFrom, strWhere
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj
	'Dim lgF2By2
	Dim strStartDt, strEndDt
	Dim strYear, strMonth, strDay

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if

    lgBlnFlgChgValue = True
	'strStartDt = UniConvDateAToB(frm1.txtAccountYear,parent.gDateFormatYYYY,parent.gServerDateFormat)

	If TRim(frm1.txtDeptCd.value) <>"" Then
		'Call parent.ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
			strStartDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtAccountYear.text, "01", "01")  
			strEndDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtAccountYear.text, "12", "31")  
			strStartDt =  UNIConvDateToYYYYMMDD(strStartDt, gDateFormat,Parent.gServerDateType)  
			strEndDt =  UNIConvDateToYYYYMMDD(strEndDt, gDateFormat,Parent.gServerDateType)  
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strStartDt , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strEndDt , "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

			'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus

		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
			jj = Ubound(arrVal1,1)

			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next
		End If
	End IF

End Sub



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
    gMouseClickStatus = "SPC"	'Split 상태코드 
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

	If Row < 1 Then Exit Sub
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
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
	'	If lgStrPreglno <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If

End Sub

'========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

 dim i
 dim RowList
 Dim intRetCD
    If Row <> NewRow And NewRow > 0 Then
	CALL vspdData_Click(1, NewRow)
	Set gActiveElement = document.activeElement 
    End If
 
End Sub

'========================================================================================================
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtFromEnterDt.Action = 7
       Call SetFocusToDocument("M")
       frm1.fpdtFromEnterDt.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : fpdtToEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtToEnterDt.Action = 7
       Call SetFocusToDocument("M")
       frm1.fpdtToEnterDt.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
	   Call MainQuery()
	End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
	   Call MainQuery()
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
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere
	Dim IntRetCD
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim strStartDt, strEndDt
	Dim strYear, strMonth, strDay

	CheckOrgChangeId = True

	With frm1

		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'Call parent.ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
			strStartDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtAccountYear.text, "01", "01")
			strEndDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtAccountYear.text, "12", "31")
			strStartDt =  UNIConvDateToYYYYMMDD(strStartDt, gDateFormat,Parent.gServerDateType)
			strEndDt =  UNIConvDateToYYYYMMDD(strEndDt, gDateFormat,Parent.gServerDateType)
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strStartDt , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strEndDt , "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With
		'----------------------------------------------------------------------------------------

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><IMG src="../../image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기초분개장조회</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><IMG src="../../image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>회계년도</TD>
						            <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5131ma1_fpDateTime1_txtAccountYear.js'></script></TD>
						            <TD CLASS=TD5 NOWRAP>부서코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtdeptcd" ALT="부서코드" Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag="1NXXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept()">
														 <INPUT NAME="txtdeptnm" ALT="부서명"   Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag="14N"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBizArea"   ALT="사업장"   Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: left" tag   ="1NXXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizArea.Value, 1)">
														 <INPUT NAME="txtBizAreaNm" ALT="사업장명" Size="24" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag   ="14N"></TD>
								<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtGlInputType" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="전표입력경로코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGlInputType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtGlInputType.value,'3')"> <INPUT TYPE="Text" NAME="txtGLInputTypeNm" SIZE=18 tag="14X" ALT="전표입력경로명"></TD>
								</TR>
								 <TR>
									<TD CLASS=TD5 NOWRAP>코스트센타</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCOST_CENTER_CD" MAXLENGTH="10" SIZE=12 ALT ="코스트센타 코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="Call OpenPopup(frm1.txtCOST_CENTER_CD.value, 2)">
														 <INPUT NAME="txtCOST_CENTER_NM" MAXLENGTH="20" SIZE=24 STYLE="TEXT-ALIGN:left" ALT ="코스트센타명" tag="14"></TD>
<!--									<TD CLASS=TD5 NOWRAP>금액</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5131ma1_OBJECT1_txtAmtFr.js'></script>&nbsp;~&nbsp;
										 <script language =javascript src='./js/a5131ma1_OBJECT2_txtAmtTo.js'></script>
									</TD>
-->
									<TD CLASS=TD5 NOWRAP>참조번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호" MAXLENGTH="30" SIZE="20" tag="11XXXU" ></TD>
								</TR>
								 <TR>
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="35" tag="11" ></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN ="2">
									<script language =javascript src='./js/a5131ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtAccountYear" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizArea" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtdeptcd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="horgchangeid" tag="" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtglno" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtCOST_CENTER_CD" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtRefNo" tag="" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtGlInputType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtDesc" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
