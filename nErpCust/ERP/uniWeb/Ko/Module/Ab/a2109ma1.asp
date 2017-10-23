<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Basic Info.
'*  3. Program ID           : A2109MA1
'*  4. Program Name         : 신용카드등록 
'*  5. Program Desc         : Register of Credit Card
'*  6. Component List       : B14011, B14018
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2001/02/14
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Song, Mun Gil / Cho Ig Sung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "a2109mb1.asp"			'☆: 비지니스 로직 ASP명 

'==========================================================================================================
'⊙: Grid Columns

Dim C_CREDIT_NO
Dim C_CREDIT_NM
Dim C_CREDIT_ENG_NM
Dim	C_CARD_CO_CD
Dim	C_CARD_CO_PB
Dim	C_CARD_CO_NM
Dim C_CREDIT_TYPE_CD
Dim C_CREDIT_TYPE_NM
Dim C_COST_CD
Dim C_COST_PB
Dim C_COST_NM
Dim C_RGST_NO
Dim C_EXPIRE_DT
Dim C_STTL_DT
Dim C_USE_ID
Dim C_USE_ID_NM
Dim C_USE_ID_PB
Dim C_BANK_CD
Dim C_BANK_PB
Dim C_BANK_NM
Dim C_BANK_ACCT_NO
Dim C_BANK_ACCT_PB


Sub InitSpreadPosVariables()
	C_CREDIT_NO      = 1
	C_CREDIT_NM      = 2
	C_CREDIT_ENG_NM  = 3
	C_CARD_CO_CD	 = 4
	C_CARD_CO_PB	 = 5
	C_CARD_CO_NM	 = 6
	C_CREDIT_TYPE_CD = 7
	C_CREDIT_TYPE_NM = 8
	C_COST_CD        = 9
	C_COST_PB        = 10
	C_COST_NM        = 11
	C_RGST_NO        = 12
	C_EXPIRE_DT      = 13
	C_STTL_DT        = 14
	C_USE_ID         = 15
	C_USE_ID_NM		 = 16
	C_USE_ID_PB		 = 17
	C_BANK_CD        = 18
	C_BANK_PB        = 19
	C_BANK_NM        = 20
	C_BANK_ACCT_NO   = 21
	C_BANK_ACCT_PB   = 22
End Sub

 '==========================================  1.2.2 Global 변수 선언  =====================================
 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey

<%
Dim dtToday
dtToday = GetSvrDate
%>
'========================================================================================================= 

<!-- #Include file="../../inc/lgvariables.inc" -->	

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    '---- Coding part--------------------------------------------------------------------

    lgStrPrevKey = ""
    lgLngCurRows = 0

    lgPageNo  = 0
    lgSortKey = 1
End Sub

'========================================================================================================= 

Sub SetDefaultVal()
	' 현재 Page의 Form Element들을 Clear한다. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear

End Sub


'======================================================================================== 

Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

'========================================================================================

Sub InitSpreadSheet()

	Call initSpreadPosVariables()	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread	

	With frm1.vspdData
		.ReDraw = False

		.MaxCols = C_BANK_ACCT_PB + 1
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
'		.ColsFrozen = C_CREDIT_NO

        Call AppendNumberPlace("6","2","0")

'		SSSetEdit(Col, Header, ColWidth, HAlign, Row, Length, CharCase)


		ggoSpread.SSSetEdit C_CREDIT_NO, "카드번호", 15, , , 20
		ggoSpread.SSSetEdit C_CREDIT_NM, "카드명", 20, , , 20
		ggoSpread.SSSetEdit C_CREDIT_ENG_NM , "카드영문명", 20, , , 50
		ggoSpread.SSSetEdit C_CARD_CO_CD, "카드사코드", 10, , , 10
		ggoSpread.SSSetButton C_CARD_CO_PB
		ggoSpread.SSSetEdit C_CARD_CO_NM, "카드사명", 20, , , 30
		ggoSpread.SSSetCombo C_CREDIT_TYPE_CD, "카드구분", 1
		ggoSpread.SSSetCombo C_CREDIT_TYPE_NM, "카드구분", 10
		ggoSpread.SSSetEdit C_COST_CD, "코스트센타코드", 15, , , 10, 2
		ggoSpread.SSSetButton  C_COST_PB
		ggoSpread.SSSetEdit C_COST_NM, "코스트센타명", 20, , , 20
		ggoSpread.SSSetEdit C_RGST_NO, "주민등록번호", 15, , , 20
		ggoSpread.SSSetDate C_EXPIRE_DT, "만기일", 12, 2, Parent.gDateFormat
		ggoSpread.SSSetFloat C_STTL_DT, "결재일", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, 2, , , "1", "31"
		ggoSpread.SSSetEdit C_USE_ID , "관리자", 10, , , 13, 2
		ggoSpread.SSSetEdit C_USE_ID_NM , "사용자이름", 15, , , 30, 2
		ggoSpread.SSSetButton C_USE_ID_PB
		ggoSpread.SSSetEdit C_BANK_CD, "은행코드", 10, , , 10, 2
		ggoSpread.SSSetButton  C_BANK_PB
		ggoSpread.SSSetEdit C_BANK_NM, "은행명", 20, , , 30
		ggoSpread.SSSetEdit C_BANK_ACCT_NO, "계좌번호", 20, , , 30
		ggoSpread.SSSetButton C_BANK_ACCT_PB

		call ggoSpread.MakePairsColumn(C_CREDIT_TYPE_CD,C_CREDIT_TYPE_NM,"1")
		call ggoSpread.MakePairsColumn(C_COST_CD,C_COST_PB,"1")
		call ggoSpread.MakePairsColumn(C_CARD_CO_CD,C_CARD_CO_NM,"1")
		call ggoSpread.MakePairsColumn(C_USE_ID,C_USE_ID_PB,"1")
		call ggoSpread.MakePairsColumn(C_BANK_CD,C_BANK_PB,"1")

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_USE_ID_NM,C_USE_ID_NM,True)
		Call ggoSpread.SSSetColHidden(C_CREDIT_TYPE_CD,C_CREDIT_TYPE_CD,True)

		.ReDraw = True

		Call SetSpreadLock

    End With
    
End Sub


'========================================================================================

Sub SetSpreadLock()

    With frm1.vspdData
		.ReDraw = False

		ggoSpread.SpreadLock C_CREDIT_NO, -1, C_CREDIT_NO
		ggoSpread.SSSetRequired C_CREDIT_NM,-1, C_CREDIT_NM
		ggoSpread.SpreadLock C_CARD_CO_NM, -1, C_CARD_CO_NM
		ggoSpread.SpreadLock C_COST_NM, -1, C_COST_NM
		ggoSpread.SpreadLock C_BANK_NM, -1, C_BANK_NM

		ggoSpread.SSSetRequired C_CREDIT_TYPE_NM, -1
		ggoSpread.SSSetRequired C_COST_CD, -1
		ggoSpread.SSSetRequired C_EXPIRE_DT, -1
		ggoSpread.SSSetRequired C_STTL_DT, -1
		ggoSpread.SSSetRequired C_BANK_CD, -1
		ggoSpread.SSSetRequired C_BANK_ACCT_NO, -1
		ggoSpread.SSSetProtected .MaxCols,-1,-1
		.ReDraw = True

    End With

End Sub


'========================================================================================

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1

		.vspdData.ReDraw = False

		' 필수 입력 항목으로 설정 
		' SSSetRequired(ByVal Col, ByVal Row, Optional Row2)
		ggoSpread.SSSetRequired C_CREDIT_NO, pvStartRow, pvEndRow		' 신용카드 번호 
		ggoSpread.SSSetRequired C_CREDIT_NM,	pvStartRow, pvEndRow	' 거래유형명 
		ggoSpread.SSSetRequired C_CREDIT_TYPE_NM, pvStartRow, pvEndRow	' 신용카드 타입 
		'ggoSpread.SpreadLock	C_CARD_CO_NM,		pvStartRow, pvEndRow				' 
		ggoSpread.SSSetRequired C_EXPIRE_DT, pvStartRow, pvEndRow		' 만기일 
		ggoSpread.SSSetRequired C_STTL_DT, pvStartRow, pvEndRow			' 결재일 
		ggoSpread.SSSetRequired C_COST_CD, pvStartRow, pvEndRow			' Cost Nm
		ggoSpread.SSSetRequired C_BANK_CD, pvStartRow, pvEndRow			' Cost Nm
		ggoSpread.SSSetRequired C_BANK_ACCT_NO, pvStartRow, pvEndRow	' Bank Acct No

		ggoSpread.SSSetProtected C_CARD_CO_NM, pvStartRow, pvEndRow				' 
		ggoSpread.SSSetProtected C_COST_NM,    pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BANK_NM,    pvStartRow, pvEndRow

		.vspdData.ReDraw = True

    End With

End Sub

'========================================================================================================= 
Sub InitComboBox()

Dim iCodeArr,iNameArr

	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1016", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	iCodeArr = lgF0
    iNameArr = lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_CREDIT_TYPE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_CREDIT_TYPE_NM
End Sub


Function InitCombo(Byval strMajorCd)
End Function

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows

			.Row = intRow

			.Col = C_CREDIT_TYPE_CD
			intIndex = .value
			.col = C_CREDIT_TYPE_NM
			.value = intindex

		Next	
	End With
End Sub



'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

               	C_CREDIT_NO      = iCurColumnPos(1)
				C_CREDIT_NM      = iCurColumnPos(2)
				C_CREDIT_ENG_NM  = iCurColumnPos(3)
               	C_CARD_CO_CD     = iCurColumnPos(4)
               	C_CARD_CO_PB	 = iCurColumnPos(5)
				C_CARD_CO_NM     = iCurColumnPos(6)
				C_CREDIT_TYPE_CD = iCurColumnPos(7)
				C_CREDIT_TYPE_NM = iCurColumnPos(8)
				C_COST_CD        = iCurColumnPos(9)
				C_COST_PB        = iCurColumnPos(10)
				C_COST_NM        = iCurColumnPos(11)
				C_RGST_NO        = iCurColumnPos(12)
				C_EXPIRE_DT      = iCurColumnPos(13)
				C_STTL_DT        = iCurColumnPos(14)
				C_USE_ID         = iCurColumnPos(15)
				C_USE_ID_NM      = iCurColumnPos(16)
				C_USE_ID_PB      = iCurColumnPos(17)
				C_BANK_CD        = iCurColumnPos(18)
				C_BANK_PB        = iCurColumnPos(19)
				C_BANK_NM        = iCurColumnPos(20)
				C_BANK_ACCT_NO   = iCurColumnPos(21)
				C_BANK_ACCT_PB   = iCurColumnPos(22)

    End Select
End Sub


'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0
			arrParam(0) = "신용카드 팝업"					' 팝업 명칭 
			arrParam(1) = "B_CREDIT_CARD"						' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "신용카드"					' 조건필드의 라벨 명칭 

			arrField(0) = "CREDIT_NO"							' Field명(0)
			arrField(1) = "CREDIT_NM"							' Field명(1)

			arrHeader(0) = "신용카드번호"					' Header명(0)
			arrHeader(1) = "신용카드명"					' Header명(1)

		Case 1
			arrParam(0) = "은행 팝업"						' 팝업 명칭 
			arrParam(1) = "B_BANK"    							' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "은행"						' 조건필드의 라벨 명칭 

			arrField(0) = "BANK_CD"	     						' Field명(0)
			arrField(1) = "BANK_NM"				    			' Field명(1)

			arrHeader(0) = "은행코드"						' Header명(0)
			arrHeader(1) = "은행명"						' Header명(1)

		Case 2
			arrParam(0) = "계좌번호 팝업"					' 팝업 명칭 
			arrParam(1) = "B_BANK_ACCT A, B_BANK B"	    		' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD"				' Where Condition
			arrParam(5) = "계좌번호"						' 조건필드의 라벨 명칭 

			With frm1.vspdData
				.Col = C_BANK_CD
				If Trim(.Text) <> "" Then arrParam(4) = arrParam(4) & " AND A.BANK_CD =  " & FilterVar(.Text, "''", "S") & ""
			End With

			arrField(0) = "A.BANK_ACCT_NO"	    				' Field명(0)
			arrField(1) = "A.BANK_CD"
			arrHeader(0) = "계좌번호"						' Header명(0)
			arrHeader(1) = "은행코드"						' Header명(0)

		Case 3
			arrParam(0) = "코스트센타 팝업"				' 팝업 명칭 
			arrParam(1) = "B_Cost_Center" 						' TABLE 명칭 
			arrParam(2) = strCode								' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "코스트센타"					' 조건필드의 라벨 명칭 

			arrField(0) = "COST_CD"									' Field명(0)
			arrField(1) = "COST_NM"									' Field명(1)

			arrHeader(0) = "코스트센타코드"					' Header명(0)
			arrHeader(1) = "코스트센타명"					' Header명(1)
		Case 4
			arrParam(0) = "사용자이름팝업"
			arrParam(1) = "Haa010t a, b_minor b, b_minor c"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "b.major_cd = " & FilterVar("H0001", "''", "S") & " and a.pay_grd1 = b.minor_cd and c.major_cd = " & FilterVar("H0002", "''", "S") & " and a.roll_pstn = c.minor_cd and a.internal_cd LIKE " & FilterVar("%", "''", "S") & " And emp_no>= ''"

			arrParam(5) = "사원"

			arrField(0) = "A.name"
			arrField(1) = "A.emp_no"
			arrField(2) = "A.dept_nm"
			arrField(3) = "C.minor_nm"
			arrField(4) = "b.minor_nm"
			arrField(5) = "A.entr_dt"
			arrField(6) = "a.dept_cd"

			arrHeader(0) = "이름"
			arrHeader(1) = "사원번호"
			arrHeader(2) = "부서명"
			arrHeader(3) = "직위"
			arrHeader(4) = "급호"
			arrHeader(5) = "입사일"
			arrHeader(6) = "부서코드"
		Case 5
			arrParam(0) = "카드사팝업"
			arrParam(1) = "B_CARD_CO A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""

			arrParam(5) = "카드사코드"

			arrField(0) = "A.CARD_CO_CD"
			arrField(1) = "A.CARD_CO_NM"

			arrHeader(0) = "카드사코드"
			arrHeader(1) = "카드사명"

		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call GridSetFocus(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If

End Function
'=======================================================================================================
Function GridSetFocus(Byval iWhere)
		With frm1
			Select Case iWhere
			Case 0
				frm1.txtCredit_No.focus
			End Select
		End With
End Function

'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtCredit_No.value = Trim(arrRet(0))
				.txtCredit_Nm.value = arrRet(1)
			Case 1
			    .vspdData.Col = C_BANK_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_BANK_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )
			Case 2
				.vspdData.Col = C_BANK_ACCT_NO
				.vspdData.Text = arrRet(0)
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )
			Case 3
			    .vspdData.Col = C_COST_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_COST_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )
			Case 4
				.vspdData.Col = C_USE_ID
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )
			Case 5
			    .vspdData.Col = C_CARD_CO_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_CARD_CO_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )
		End Select
	End With
End Function


'========================================================================================================= 

Function OpenCalendar(Byval iWhere)
	Dim RetCal

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
	RetCal = window.showModalDialog("../../comasp/Calendar.asp", , _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	If RetCal = "" Then
		Exit Function
	Else
		Call SetCalendar(RetCal, iWhere)
	End If
End Function


'========================================================================================================= 
Function SetCalendar(Byval RetCal, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 1
				.vspdData.Col = C_EXPIRE_DT
				.vspdData.Text = RetCal
				Call vspdData_Change(.vspdData.Col,.vspdData.Row )	
		End Select
'		lgBlnFlgChgValue = True
	End With
End Function


'========================================================================================================= 
Sub Form_Load()

    On Error Resume Next
    Err.Clear

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    Call InitComboBox
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call FncSetToolBar("New")
	frm1.txtCredit_No.focus 
End Sub

'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If

End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------   
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData

		.Row = Row

		Select Case Col
			Case  C_CREDIT_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_CREDIT_TYPE_CD
				.Value = intIndex

'            Case C_CREDIT_TYPE_CD
'                .Col = Col
'                intIndex = .Value
'                .Col = C_CREDIT_TYPE_NM
'                .Value = intIndex
		End Select
	End With
End Sub

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows

			.Row = intRow

			.Col = C_CREDIT_TYPE_CD
			intIndex = .value
			.col = C_CREDIT_TYPE_NM
			.value = intindex

		Next
	End With
End Sub



'==========================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )
   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 

   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row

End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData 

		If Row >= NewRow Then
		    Exit Sub
		End If

    End With

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
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	'---------- Coding part -------------------------------------------------------------
	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			Select Case Col
			Case C_BANK_PB
				.Col = Col
				.Row = Row

				.Col = C_BANK_CD
				Call OpenPopup(.Text,1)
			Case C_BANK_ACCT_PB
				.Col = Col
				.Row = Row

				.Col = C_BANK_ACCT_NO
				Call OpenPopup(.Text,2)
			Case C_COST_PB
				.Col = Col
				.Row = Row

				.Col = C_COST_CD
				Call OpenPopup(.Text,3)
			Case C_USE_ID_PB
				.Col = Col
				.Row = Row

				.Col = C_USE_ID
				Call OpenPopup(.Text,4)
			Case C_CARD_CO_PB
				.Col = Col
				.Row = Row

				.Col = C_CARD_CO_CD
				Call OpenPopup(.Text,5)
			End Select
		End If
		Call SetActiveCell(frm1.vspdData,Col-1,frm1.vspdData.ActiveRow ,"M","X","X")
	End With
	
End Sub

'========================================================================================
Function FncQuery()
	Dim IntRetCD

    FncQuery = False

    On Error Resume Next
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
	' 현재 Page의 Form Element들을 Clear한다. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "2")
    Call InitVariables
    Call InitComboBox

    '-----------------------
    'Check condition area
    '-----------------------
	' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
	' ChkField(pDoc, pStrGrp) As Boolean
    If Not chkField(Document, "1") Then
       Exit Function
    End If

    Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery

    FncQuery = True

End Function


'========================================================================================

Function FncNew() 
	Dim IntRetCD 

    FncNew = False
    Err.Clear
    On Error Resume Next

    '-----------------------
    'Check previous data area
    '-----------------------
    ' 변경된 내용이 있는지 확인한다.
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015",Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------

    Call ggoOper.ClearField(Document, "1")
	Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call SetDefaultVal

    Call FncSetToolBar("New")
    FncNew = True

End Function


'========================================================================================

Function FncDelete()
	Dim IntRetCD 

    FncDelete = False
    Err.Clear
    On Error Resume Next

    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")

	If IntRetCD = vbNo Then
		Exit Function
	End If

    If DbDelete = False Then
       Exit Function
    End If

	Call ggoOper.ClearField(Document, "1")
    FncDelete = True
End Function


'========================================================================================
Function FncSave()
	Dim IntRetCD 

    FncSave = False
    Err.Clear
    On Error Resume Next

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData 
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    

    Call DbSave
	 FncSave = True
End Function


'========================================================================================
Function FncCopy()
	Dim  IntRetCD

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	frm1.vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

	frm1.vspdData.Col	= C_CREDIT_NO
	frm1.vspdData.text	= ""

	frm1.vspdData.Redraw = True
End Function



Function FncCancel()
	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo

    Call InitData
End Function


'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next
    Err.Clear

    FncInsertRow = False

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else
    imRow = AskSpdSheetAddRowCount()

    If imRow = "" Then
        Exit Function
		End If
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True
    End If

    Set gActiveElement = document.ActiveElement
End Function


'========================================================================================
Function FncDeleteRow()
    Dim lDelRows

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		lDelRows = ggoSpread.DeleteRow
    End With

End Function


'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function


'========================================================================================
Function DbQuery()
	Dim strVal

	Call LayerShowHide(1)

    DbQuery = False
    Err.Clear

    With frm1

		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtCredit_No=" & Trim(.htxtCredit_No.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtCredit_No=" & Trim(.txtCredit_No.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		strVal = strVal & "&lgPageNo="       & lgPageNo

	    Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 

    End With

    DbQuery = True

End Function


'========================================================================================
Function DbQueryOk()

	Call SetSpreadLock
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
    Call InitData 
	Call FncSetToolBar("Query")
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement 
End Function


'========================================================================================
Function DbSave()
	Dim lRow
	Dim lGrpCnt
	Dim strVal,strDel

	Call LayerShowHide(1)

    DbSave = False

	With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows

		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag											'☜: 신규 
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep				'☜: C=Create
		            .vspdData.Col = C_CREDIT_NO
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CREDIT_NM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CREDIT_ENG_NM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CARD_CO_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_CREDIT_TYPE_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_COST_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_RGST_NO
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            strVal = strVal & Trim(UNIConvDate("")) & Parent.gColSep
		            .vspdData.Col = C_EXPIRE_DT
		            strVal = strVal & Trim(UNIConvDate(.vspdData.Text)) & Parent.gColSep
		            .vspdData.Col = C_STTL_DT
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_USE_ID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_BANK_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_BANK_ACCT_NO
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep

		            lGrpCnt = lGrpCnt + 1

				Case ggoSpread.UpdateFlag												'☜: 수정 

					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep					'☜: U=Update
				    .vspdData.Col = C_CREDIT_NO
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CREDIT_NM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CREDIT_ENG_NM
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CARD_CO_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_CREDIT_TYPE_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_COST_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_RGST_NO
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            strVal = strVal & Trim(UNIConvDate("")) & Parent.gColSep
		            .vspdData.Col = C_EXPIRE_DT
		            strVal = strVal & Trim(UNIConvDate(.vspdData.Text)) & Parent.gColSep
		            .vspdData.Col = C_STTL_DT
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_USE_ID
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_BANK_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_BANK_ACCT_NO
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep

		            lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag												'☜: 삭제 

					strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep					'☜: U=Delete
		            .vspdData.Col = C_CREDIT_NO
		            strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep

		            lGrpCnt = lGrpCnt + 1
		    End Select

		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal

		 Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'☜: 비지니스 ASP 를 가동 

	End With

    DbSave = True
End Function


'========================================================================================
Function DbSaveOk()
	Call ggoOper.ClearField(Document, "2")
    Call InitVariables
	frm1.vspdData.MaxRows = 0
    Call FncQuery()
End Function


'========================================================================================
Function DbDelete()
	On Error Resume Next
End Function

'========================================================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100110100101111")
	Case "QUERY"
		Call SetToolbar("1100111100111111")
	End Select
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>신용카드등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>신용카드번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtCredit_No" MAXLENGTH="20" SIZE=20 ALT ="신용카드 번호" tag="11X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtCredit_No.Value,0)">
														   <INPUT NAME="txtCredit_Nm" MAXLENGTH="20" SIZE=30 STYLE="TEXT-ALIGN:left" ALT ="신용카드명" tag="24X"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/a2109ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" src="../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME></TD>
		
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="htxtCredit_No" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
