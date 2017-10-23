<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        :
'*  3. Program ID           : A5967MA1
'*  4. Program Name         : 상여금 실지급액 등록 
'*  5. Program Desc         : 상여금 실지급액 등록 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : 권기수 
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'========================================================================================================
Const BIZ_PGM_ID = "A5967MB1.asp"                                      'Biz Logic ASP
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Dim C_DEPT_CD                                                 'Column constant for Spread Sheet
Dim C_DEPT_CD_H
Dim C_DEPT_CD_BT
Dim C_DEPT_NM 
Dim C_BIZ_AREA_CD
Dim C_ORG_CHANGE_ID
Dim C_INTERNAL_CD
Dim C_ACCT_TYPE  
Dim C_ACCT_TYPE_CD
Dim C_ACCT_TYPE_CD_H
Dim C_PAY_AMT 


Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String
'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop
Dim IsOpenPop
Dim lsConcd

'========================================================================================================
Sub initSpreadPosVariables()
	 C_DEPT_CD         = 1                                                 'Column constant for Spread Sheet
	 C_DEPT_CD_H       = 2
	 C_DEPT_CD_BT      = 3
	 C_DEPT_NM         = 4
	 C_BIZ_AREA_CD     = 5
	 C_ORG_CHANGE_ID   = 6
	 C_INTERNAL_CD		= 7
	 C_ACCT_TYPE       = 8
	 C_ACCT_TYPE_CD    = 9
	 C_ACCT_TYPE_CD_H  = 10
	 C_PAY_AMT      = 11
End Sub
'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
	
	frm1.fpdtWk_yyyy.Year	= strYear
	frm1.fpdtWk_yyyy.Month	= strMonth
	frm1.fpdtWk_yyyy.Day	= strDay
	
	Call ggoOper.FormatDate(frm1.fpdtWk_yyyy, gDateFormat, 3)
	
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) --------------------------------------------------------------
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : Make key stream of query or delete condition data
'========================================================================================================
Sub MakeKeyStream(pOpt)

    Dim strYYYY
    Dim strYear,strMonth,strDay

   '------ Developer Coding part (Start ) --------------------------------------------------------------

    Call ExtractDateFrom(frm1.fpdtWk_yyyy.text,frm1.fpdtWk_yyyy.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

    'strYYYY = strYear
    strYYYY = Trim(frm1.fpdtWk_yyyy.text)
    lgKeyStream = strYYYY & parent.gColSep       'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Trim(frm1.txtBonusCd.value) & parent.gColSep                      '상여종류 
    lgKeyStream = lgKeyStream & Trim(frm1.txtBizAreaCd.value) & parent.gColSep                      '상여종류 
    lgKeyStream = lgKeyStream & Trim(frm1.selPayMm.value) & parent.gColSep                      '상여종류 
   '------ Developer Coding part (End   ) --------------------------------------------------------------

End Sub


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0071", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ACCT_TYPE
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ACCT_TYPE_CD
	'------ Developer Coding part (End )   --------------------------------------------------------------
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
			.Col = C_ACCT_TYPE
			intIndex = .value
			.col = C_ACCT_TYPE_CD
			.value = intindex
		Next
	End With
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
		.ReDraw = false
		
       .MaxCols = C_PAY_AMT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:


		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit    C_DEPT_CD         , "부서코드"   ,17,,, 35,2
		ggoSpread.SSSetEdit    C_DEPT_CD_H       , "부서코드"   ,15,,, 35,2
		ggoSpread.SSSetButton  C_DEPT_CD_BT
		ggoSpread.SSSetEdit    C_DEPT_NM         , "부서명"     ,40,,, 35,2
		ggoSpread.SSSetEdit    C_BIZ_AREA_CD     , "사업장"   ,20,,, 35,2
		ggoSpread.SSSetEdit    C_ORG_CHANGE_ID   , "조직변경ID"   ,20,,, 35,2
		ggoSpread.SSSetEdit    C_INTERNAL_CD		, "내부부서코드"   ,20,,, 35,2
		ggoSpread.SSSetCombo   C_ACCT_TYPE       , "계정TYPE"   ,17
		ggoSpread.SSSetCombo   C_ACCT_TYPE_CD    , "계정TYPE"   ,05
		ggoSpread.SSSetEdit    C_ACCT_TYPE_CD_H  , "계정TYPE"   ,05,,, 35,2
		ggoSpread.SSSetFloat   C_PAY_AMT         , "실지급액"     ,35, parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

		Call ggoSpread.MakePairsColumn(C_DEPT_CD_H,C_DEPT_CD_BT,"1")
		Call ggoSpread.SSSetColHidden(C_DEPT_CD_H,C_DEPT_CD_H,True)
		Call ggoSpread.SSSetColHidden(C_BIZ_AREA_CD,C_BIZ_AREA_CD,True)
		Call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID,C_ORG_CHANGE_ID,True)
		Call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_INTERNAL_CD,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_TYPE_CD,C_ACCT_TYPE_CD,True)
		Call ggoSpread.SSSetColHidden(C_ACCT_TYPE_CD_H,C_ACCT_TYPE_CD_H,True)
		
	   .ReDraw = true

       Call SetSpreadLock
    End With
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired C_DEPT_CD, -1 , C_DEPT_CD
        ggoSpread.SpreadLock    C_DEPT_NM, -1 , C_DEPT_NM
        ggoSpread.SSSetRequired    C_ACCT_TYPE , -1 , C_ACCT_TYPE
        ggoSpread.SpreadLock    C_ACCT_TYPE_CD , -1 , C_ACCT_TYPE_CD
        ggoSpread.SSSetRequired    C_PAY_AMT , -1, C_PAY_AMT
        ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols
        .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
            ggoSpread.SSSetRequired    C_DEPT_CD , pvStarRow, pvEndRow
            ggoSpread.SSSetProtected   C_DEPT_NM , pvStarRow, pvEndRow
            ggoSpread.SSSetRequired    C_ACCT_TYPE , pvStarRow, pvEndRow
            ggoSpread.SSSetRequired    C_PAY_AMT, pvStarRow, pvEndRow
        .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_DEPT_CD         = iCurColumnPos(1)                                                 'Column constant for Spread Sheet
		C_DEPT_CD_H       = iCurColumnPos(2)
		C_DEPT_CD_BT      = iCurColumnPos(3)
		C_DEPT_NM         = iCurColumnPos(4)
		C_BIZ_AREA_CD     = iCurColumnPos(5)
		C_ORG_CHANGE_ID   = iCurColumnPos(6)
		C_INTERNAL_CD		= iCurColumnPos(7)
		C_ACCT_TYPE       = iCurColumnPos(8)
		C_ACCT_TYPE_CD    = iCurColumnPos(9)
		C_ACCT_TYPE_CD_H  = iCurColumnPos(10)
		C_PAY_AMT      = iCurColumnPos(11)
	End Select
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If
       Next
    End If
End Sub

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'======================================================================================================
Function OpenBonus()
	Dim arrRet
	Dim IntRetCD
	Dim arrParam(6), arrField(5), arrHeader(5)


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "상여종류 팝업"		    	    <%' 팝업 명칭 %>
	arrParam(1) = "a_bonus_base a,b_minor b" 		<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtBonusCd.value				<%' Code Cindition%>
	arrParam(3) = ""								<%' Name Condition%>
	If Trim(frm1.fpdtWk_yyyy.Text) <> "" Then
		arrParam(4) = "b.major_cd = " & FilterVar("H0040", "''", "S") & "  AND a.pay_type = b.minor_cd and a.yyyy = " & FilterVar(frm1.fpdtWk_yyyy.Text, "''", "S")
	Else
		IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
		IsOpenPop = False
        Exit Function
    End If    	
	arrParam(5) = "상여코드"

    arrField(0) = "b.minor_cd"					<%' Field명(0)%>
    arrField(1) = "b.minor_nm"	     			<%' Field명(1)%>

    arrHeader(0) = "상여코드"				<%' Header명(0)%>
    arrHeader(1) = "상여명"					<%' Header명(1)%>


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBonusCd.focus
		Exit Function
	Else
		Call SetBonus(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetBonus()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetBonus(Byval arrRet)
	With frm1
		.txtBonusCd.focus
		.txtBonusCd.value = arrRet(0)
		.txtBonus.value	   = arrRet(1)
	End With
End Function

'======================================================================================================
'	Name : OpenBizArea)
'	Description : Major PopUp
'======================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"		    	<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA"					<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtBizAreaCd.value       <%' Code Condition%>
	arrParam(3) = "" 		            		<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "사업장"

    arrField(0) = "BIZ_AREA_CD"	     			<%' Field명(1)%>
    arrField(1) = "BIZ_AREA_NM"					<%' Field명(0)%>


    arrHeader(0) = "사업장코드"			    <%' Header명(0)%>
    arrHeader(1) = "사업장명"				<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=470px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetBizArea(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetBizArea()
'	Description : Item Popup에서 Return되는 값 setting
'======================================================================================================
Function SetBizArea(Byval arrRet)
	With frm1
		.txtBizAreaCd.focus
		.txtBizAreaCd.value = arrRet(0)
		.txtBizArea.value	   = arrRet(1)
	End With
End Function

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row
		Select Case Col
		Case C_ACCT_TYPE
			.Col = Col
			intIndex = .Value
			.Col = C_ACCT_TYPE_CD
			.Value = intIndex
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
		End Select
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    Dim BizAreaCd
    BizAreaCd = Trim(frm1.txtBizAreaCd.value)
	Select Case Col				'추가부분을 위해..select로..
	    Case C_DEPT_CD_BT        'Cost center
	        frm1.vspdData.Col = C_DEPT_CD
	        If BizAreaCd = "" then
	            Call OpenDept(frm1.vspdData.Text,1, Row)
	        Else
	            Call OpenDept(frm1.vspdData.Text,2, Row )
	        End If
	End Select

End Sub

'======================================================================================================
'	Name : OpenDept
'	Description : 
'======================================================================================================
Function OpenDept(Byval strCode, iWhere, Row)
	Dim arrRet
	Dim arrParam(5)
	Dim strYear, strMonth, strDay, strDate
	'------ Developer Coding part (Start ) --------------------------------------------------------------
		
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode									'  Code Condition
	
	strDate = FilterVar(frm1.fpdtWk_yyyy.text,"2999","SNM") & "-" & FilterVar(frm1.selPayMm.value,"12","SNM") & "-" & "01"
	strDate = DateAdd("D",-1, DateAdd("M",1,cdate(strDate)))
   	arrParam(1) = UNIDateClientFormat(strDate)
	arrParam(2) = lgUsrIntCd								' 자료권한 Condition  

	'' T : protected F: 필수 
	'If lgIntFlgMode = parent.OPMD_UMODE then
		arrParam(3) = "T"									' 결의일자 상태 Condition  
	'Else
	'	arrParam(3) = "F"									' 결의일자 상태 Condition  
	'End If
	
	arrParam(4) = iWhere
	arrParam(5) = Trim(frm1.txtBizAreaCd.value)
	
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDt3.asp", Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
				Call SetActiveCell(frm1.vspdData,C_DEPT_CD,frm1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	Else
		Call SetDept(arrRet, iWhere, Row)
	End If	
			
End Function
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere, Byval iRow)
		
	With frm1
		.vspdData.Row = iRow
		Select Case iWhere
		     Case 1,2
                .vspdData.Col = C_DEPT_CD
				.vspdData.text = arrRet(0)
				.vspdData.Col = C_DEPT_NM
				.vspdData.text = arrRet(1)
				.vspdData.Col = C_BIZ_AREA_CD
				.vspdData.text = arrRet(2)
				.vspdData.Col = C_ORG_CHANGE_ID
				.vspdData.text = arrRet(3)
				.vspdData.Col = C_INTERNAL_CD	
				.vspdData.text = arrRet(4)
        End Select
	    Call vspdData_Change(C_DEPT_CD,iRow)
		Call SetActiveCell(.vspdData,C_DEPT_CD,.vspdData.ActiveRow ,"M","X","X")
	End With
End Function       
'======================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field

    Call InitSpreadSheet
	Call InitVariables
    Call SetDefaultVal
    Call InitComboBox

	Call SetToolbar("1100110100101111")
    Call BtnDisabled(1)
	Call CookiePage (0)
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'======================================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False															 '
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Call MakeKeyStream("X")
	'------ Developer Coding part (End )   --------------------------------------------------------------

    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncQuery = True

End Function

'======================================================================================================
Function FncNew()

End Function

'======================================================================================================
Function FncDelete()
    Dim intRetCD

    FncDelete = False                                                            '
    Err.Clear

   If lgIntFlgMode <> parent.OPMD_UMODE Then                                           'Check if there is retrived data
   Call DisplayMsgBox("900002","X","X","X")                                 '☜: Please do Display first.
   Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		                 '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If


    Set gActiveElement = document.ActiveElement
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD

    FncSave = False                                                              '
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data.
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData


	If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If


    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call MakeKeyStream("X")
	'------ Developer Coding part (End )   --------------------------------------------------------------
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement
    FncSave = True                                                              '☜: Processing is OK
End Function

'======================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '
    Err.Clear

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow

            .ReDraw = True
		    .Focus
		 End If
	End With

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	' Clear key field
	'----------------------------------------------------------------------------------------------------
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.col = C_DEPT_CD 
			frm1.vspdData.text = ""
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.col = C_DEPT_NM 
			frm1.vspdData.text = ""
			
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCopy = True
End Function

'======================================================================================================
Function FncCancel()
    FncCancel = False                                                            '
    Err.Clear

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncCancel = False                                                            '☜: Processing is OK
End Function

'======================================================================================================
Function FncInsertRow(Byval pvRowCnt)
	Dim imRow
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If
	
	If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If


	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow

        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'======================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    FncDeleteRow = False                                                         '
    Err.Clear

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if

    With Frm1.vspdData
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement
    FncDeleteRow = True
                                                 '☜: Processing is OK
End Function

'======================================================================================================
Function FncPrint()
    FncPrint = False
    Err.Clear
	Call Parent.FncPrint()
    FncPrint = True
End Function

'======================================================================================================
Function FncPrev()
    FncPrev = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncPrev = True
End Function

'======================================================================================================
Function FncNext()
    FncNext = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncNext = True
End Function

'========================================================================================================
Function FncExcel()
    FncExcel = False
    Err.Clear

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True
End Function

'========================================================================================================
Function FncFind()
    FncFind = False
    Err.Clear
	Call Parent.FncFind(parent.C_MULTI, True)
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

'========================================================================================================
Function FncExit()
	Dim IntRetCD
    FncExit = False
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear
    DbQuery = False                                                              '

    if LayerShowHide(1) = false then                                                        '☜: Show Processing Message
		exit function
	end if

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '☜: Next key tag
    End With

    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function DbSave()
    Dim pP21011
    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
	Dim strVal
	Dim strDel

    Err.Clear

    DbSave = False                                                               '
   if LayerShowHide(1) = false then                                                        '☜: Show Processing Message
		exit function
	end if
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
  	With frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    lGrpCnt = 1

  	With frm1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

            Case ggoSpread.InsertFlag                                      '☜: Update
                                                      strVal = strVal & "C" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                .vspdData.Col = C_DEPT_CD	        : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_BIZ_AREA_CD	    : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_ORG_CHANGE_ID     : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_INTERNAL_CD		: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_ACCT_TYPE_CD	    : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_PAY_AMT	        : strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.UpdateFlag                                      '☜: Update
                                                      strVal = strVal & "U" & parent.gColSep
                                                      strVal = strVal & lRow & parent.gColSep
                .vspdData.Col = C_DEPT_CD	        : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_BIZ_AREA_CD	    : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_ORG_CHANGE_ID     : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_INTERNAL_CD		: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_ACCT_TYPE_CD	    : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_PAY_AMT	        : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_DEPT_CD_H         : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_ACCT_TYPE_CD_H    : strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                lGrpCnt = lGrpCnt + 1

            Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                      strDel = strDel & "D" & parent.gColSep
                                                      strDel = strDel & lRow & parent.gColSep
                .vspdData.Col = C_DEPT_CD_H	        : strDel = strDel & Trim(.vspdData.text) & parent.gColSep
                .vspdData.Col = C_ACCT_TYPE_CD_H	: strDel = strDel & Trim(.vspdData.text) & parent.gRowSep
                lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
  	   .txtMaxRows.value     = lGrpCnt-1
	   .txtSpread.value      =  strDel & strVal

	End With


	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear
    DbDelete = False
    if LayerShowHide(1) = false then
		exit function
	end if

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    DbDelete = True
    Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================================
Sub DbQueryOk()

    lgIntFlgMode = parent.OPMD_UMODE

	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
         Call BtnDisabled(1)
    end if
   
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub


'========================================================================================================
Sub DbSaveOk()
    Call InitVariables
	'------ Developer Coding part (Start)  --------------------------------------------------------------
    Call MakeKeyStream("X")
     Call ggoOper.ClearField(Document, "2")									     '⊙: Clear Contents  Field
     ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	'------ Developer Coding part (End )   --------------------------------------------------------------

    Call DisableToolBar(parent.TBC_QUERY)
    If DBQuery = false Then
        Call RestoreToolBar()
        Exit Sub
    End If

	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if
   Set gActiveElement = document.ActiveElement
End Sub

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	Call InitVariables()
	If frm1.vspdData.MaxRows = 0 then
       Call SetToolbar("1100110100101111")
    else
       Call SetToolbar("1100111100111111")
    end if
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,Input_alloc,  EFlag
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	Dim strDate
	EFlag = False

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) --------------------------------------------------------------
    Select Case Col
		Case C_DEPT_CD 
			If Trim(frm1.txtBizAreaCd.value) = "" then
				Call DisplayMsgBox("169803","X","X","X") 
				frm1.vspdData.Text=""
				frm1.vspdData.Col = C_DEPT_NM
				frm1.vspdData.Text=""
				frm1.vspdData.Col = C_BIZ_AREA_CD
				frm1.vspdData.Text=""
				frm1.vspdData.Col = C_ORG_CHANGE_ID
				frm1.vspdData.Text=""
				frm1.txtBizAreaCd.focus
				Set gActiveElement = document.activeElement  
				EFlag = True
			Else
				If Trim(frm1.vspdData.text) = "" Then
					Frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_DEPT_NM
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_BIZ_AREA_CD
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_ORG_CHANGE_ID
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_INTERNAL_CD
					frm1.vspdData.Text=""
					ggoSpread.Source = frm1.vspdData
					ggoSpread.UpdateRow Row
					Exit Sub
				End If
				
				strDate = FilterVar(frm1.fpdtWk_yyyy.text,"2999","SNM") & "-" & FilterVar(frm1.selPayMm.value,"12","SNM") & "-" & "01"
				strDate = DateAdd("D",-1, DateAdd("M",1,cdate(strDate)))
				'strDate = UNIConvDate(strDate)
				
			
				frm1.vspdData.Col = C_DEPT_CD
				Frm1.vspdData.Row = Row
				
				strSelect	=			 " a.dept_cd,a.dept_nm, a.org_change_id, a.internal_cd, b.biz_area_cd "    		
				strFrom		=			 " b_acct_dept a, b_cost_center b "		
				strWhere	= " a.cost_cd = b.cost_cd " 	 
				strWhere	= strWhere & " and a.dept_Cd = " & FilterVar(LTrim(RTrim(frm1.vspdData.Text)), "''", "S")
				strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(strDate, "''", "S") & "))"			
						'msgbox "select " & strSelect & " from  " & strFrom & " where  " & strWhere
				If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
					IntRetCD = DisplayMsgBox("124600","X","X","X")  
					frm1.vspdData.Col = C_DEPT_NM
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_BIZ_AREA_CD
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_ORG_CHANGE_ID
					frm1.vspdData.Text=""
					frm1.vspdData.Col = C_INTERNAL_CD
					frm1.vspdData.Text=""
					Set gActiveElement = document.activeElement  
										    
				Else 
					
					arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
					jj = Ubound(arrVal1,1)
								
					For ii = 0 to jj - 1
						arrVal2 = Split(arrVal1(ii), chr(11))			
						frm1.vspdData.Col = C_DEPT_NM
						frm1.vspdData.Text=Trim(arrVal2(2))
						frm1.vspdData.Col = C_BIZ_AREA_CD
						frm1.vspdData.Text=Trim(arrVal2(5))
						frm1.vspdData.Col = C_ORG_CHANGE_ID
						frm1.vspdData.Text=Trim(arrVal2(3))
						frm1.vspdData.Col = C_INTERNAL_CD
						frm1.vspdData.Text=Trim(arrVal2(4))
					Next	
				End If
			End If		
    End Select
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)


	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0

    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	gMouseClickStatus = "SPC"
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevKeyIndex <> "" Then
      	   Call DisableToolBar(parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
        End If
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

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

Sub fpdtWk_yyyy_DblClick(Button)
	If Button = 1 Then
		frm1.fpdtWk_yyyy.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpdtWk_yyyy.Focus
	End If
End Sub

Sub fpdtWk_yyyy_change()
	If Trim(frm1.txtBonusCd.value) <> "" Then	
		frm1.txtBonusCd.value = ""
		frm1.txtBonus.value = ""
	End IF	 
End Sub
'========================================================================================================
' Name : txtBonusCd_Onchange
' Desc : developer describe this line
'========================================================================================================
Function txtBonusCd_Onchange()
    Dim IntRetCd
  
    If  frm1.txtBonusCd.value = "" Then
		frm1.txtBonus.value=""
    Else
		If Trim(frm1.fpdtWk_yyyy.Text) <> "" Then
			IntRetCD= CommonQueryRs(" b.minor_Nm "," a_bonus_base a,b_minor b ","  b.major_cd = " & FilterVar("H0040", "''", "S") & "  AND a.pay_type = b.minor_cd and a.yyyy = " & FilterVar(frm1.fpdtWk_yyyy.Text, "''", "S") & " AND  a.Pay_type = " & FilterVar(frm1.txtBonusCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
				If IntRetCD=False Then
				    Call DisplayMsgBox("800316","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
				    frm1.txtBonusCd.value=""
				    frm1.txtBonus.value=""
				    frm1.txtBonusCd.focus
				    Set gActiveElement = document.activeElement
				Else
				    frm1.txtBonus.value=Trim(Replace(lgF0,Chr(11),""))
				End If	
		Else
			IntRetCD = DisplayMsgBox("800211","x","x","x")                           '☜:There is no changed data.	
			frm1.fpdtWk_yyyy.Focus
			Set gActiveElement = document.activeElement  
		End If			    
    End if
 lgBlnFlgChgValue = True   
End Function 
'=======================================================================================================
'   Event Name : fpdtWk_yyyy_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub fpdtWk_yyyy_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>



<BODY TABINDEX="-1" SCROLL="No">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>상여금실지급액등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>년도</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/a5967ma1_fpDateTime3_fpdtWk_yyyy.js'></script>
									</TD>
									<TD NOWRAP CLASS="TD5">상여종류</TD>
									<TD NOWRAP CLASS="TD6">
										<INPUT TYPE=TEXT   NAME="txtBonusCd" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="상여종류" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBonus" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBonus()">
                                        <INPUT TYPE=TEXT   NAME="txtBonus" TAG="14XXU" SIZE=22 MAXLENGTH"50">
									</TD>
                                </TR>
                                <TR>
									<TD NOWRAP CLASS="TD5">지급월</TD>
									<TD NOWRAP CLASS="TD6"><SELECT NAME="selPayMm" TAG="12XXXU" ALT="지급월">
									                        <OPTION VALUE="01">01</OPTION>
									                        <OPTION VALUE="02">02</OPTION>
									                        <OPTION VALUE="03">03</OPTION>
									                        <OPTION VALUE="04">04</OPTION>
									                        <OPTION VALUE="05">05</OPTION>
									                        <OPTION VALUE="06">06</OPTION>
									                        <OPTION VALUE="07">07</OPTION>
									                        <OPTION VALUE="08">08</OPTION>
									                        <OPTION VALUE="09">09</OPTION>
									                        <OPTION VALUE="10">10</OPTION>
									                        <OPTION VALUE="11">11</OPTION>
									                        <OPTION VALUE="12">12</OPTION>
									                       </SELECT>
									</TD>
									<TD NOWRAP CLASS="TD5">사업장</TD>
									<TD NOWRAP CLASS="TD6">
										<INPUT TYPE=TEXT   NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="사업장" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizArea" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea()">
                                        <INPUT TYPE=TEXT   NAME="txtBizArea" TAG="14XXU" SIZE=22 MAXLENGTH"50">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_30%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/a5967ma1_OBJECT1_vspdData.js'></script>
								    </TD>
								</TR>
				                <TR>
					                <TD HEIGHT=20 WIDTH=100%>
						            <FIELDSET CLASS="CLSFLD">
							        <TABLE  CLASS="BasicTB" CELLSPACING=0>
								        <TR>
							    	        <TD CLASS=TDT NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							    	        <TD CLASS=TD5 NOWRAP>합계</TD>
							    	        <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5967ma1_txtPayAmt_txtPayAmt.js'></script></TD>
								        </TR>
							        </TABLE>
						            </FIELDSET>
					                </TD>
				                </TR>
							</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD NOWRAP WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"   TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"  TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"     TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

