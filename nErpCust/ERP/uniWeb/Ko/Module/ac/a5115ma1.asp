<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : A5115MA1
'*  4. Program Name         : 총계정원장조회 
'*  5. Program Desc         : Query of General Ledger
'*  6. Component List       : ADO
'*  7. Modified date(First) : 2001.12.26
'*  8. Modified date(Last)  : 2004/01/14
'*  9. Modifier (First)     : Chang, Sung Hee
'* 10. Modifier (Last)      : Kim Chang Jin
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

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop                                               '☜: Popup status                           
Dim lgIsOpenPop
Dim lgMark                                                  '☜: 마크                                  

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "A5115MB1.asp"
'Dim lsPoNo								                       '☆: Jump시 Cookie로 보낼 Grid value
Const C_MaxKey          = 2                                    '☆☆☆☆: Max key value
'========================================================================================
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

'========================================================================================
Sub SetDefaultVal()
	

'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------


	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	EndDate = "<%=GetSvrDate%>"
	
	Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)


	frm1.txtFromGlDt.Text = StartDate
	frm1.txtToGlDt.Text = EndDate

'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

End Sub

'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "COOKIE", "QA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("A5115MA1","S","A","V20021220",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock  
End Sub



'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'========================================================================================
Sub InitComboBox()

End Sub
 


'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	'frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0, 3
			If frm1.txtBizAreaCd.className = Parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = frm1.txtBizAreaCd.Alt
			arrParam(1) = "B_bIZ_AREA A"
			arrParam(2) = strCode
			arrParam(3) = ""
			
			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = frm1.txtBizAreaCd.Alt
	
			arrField(0) = "A.BIZ_AREA_CD"
			arrField(1) = "A.BIZ_AREA_NM"

			arrHeader(0) = frm1.txtBizAreaCd.Alt
			arrHeader(1) = frm1.txtBizAreaNm.Alt
		
		Case 1
			arrParam(0) = "계정 팝업"									' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"					' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)
	
		Case 2
			arrParam(0) = "계정 팝업"									' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"					' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)

	End Select
    
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		Case 0					'부서 
			frm1.txtBizAreaCd.focus
		Case 1
			frm1.txtFrAcctCd.focus
		Case 2
			frm1.txtToAcctCd.focus
		Case 3					'부서 
			frm1.txtBizAreaCd1.focus
		End Select
		Exit Function
	Else
		Select Case iWhere
		Case 0					'부서 
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 1
			frm1.txtFrAcctCd.focus
			frm1.txtFrAcctCd.Value	= arrRet(0)
			frm1.txtFrAcctNm.Value	= arrRet(1)			
			Call txtFrAcctCd_OnChange()						
		Case 2
			frm1.txtToAcctCd.focus
			frm1.txtToAcctCd.Value	= arrRet(0)
			frm1.txtToAcctNm.Value	= arrRet(1)	
		Case 3					'부서 
			frm1.txtBizAreaCd1.focus
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)				
		End Select
	End If	

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

		arrVal = Split(strTemp, Parent.gRowSep)

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

		Call FncQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

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
End Sub

'========================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

	Call InitVariables													'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call InitComboBox()
	Call FncSetToolBar("New")

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
Sub txtFromGlDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromGlDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtFromGlDt.Focus
	End if
End Sub

'========================================================================================
Sub txtToGlDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToGlDt.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtToGlDt.Focus
	End if
End Sub

'========================================================================================
Sub txtFromGlDt_Keypress(Key)
    If Key = 13 Then
		frm1.txtToGlDt.focus
        FncQuery()
    End If
End Sub

'========================================================================================
Sub txtToGlDt_Keypress(Key)
    If Key = 13 Then
		frm1.txtFromGlDt.focus
        FncQuery()
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
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub

'========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
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

	If CompareDateByFormat(frm1.txtFromGlDt.Text, frm1.txtToGlDt.Text, frm1.txtFromGlDt.Alt, frm1.txtToGlDt.Alt, _
						"970025", frm1.txtFromGlDt.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtFromGlDt.focus											'⊙: GL Date Compare Common Function
			Exit Function
	End if
   
	IF frm1.txtFrAcctCd.value > frm1.txtToAcctCd.value then
		Call DisplayMsgBox("970025", "X", frm1.txtFrAcctCd.Alt, frm1.txtToAcctCd.Alt)		
		frm1.txtFrAcctCd.focus
		Exit Function
	END IF
	
	'-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	

	Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    IF  DbQuery	= False Then														'☜: Query db data
		Exit Function
	END IF
	
    FncQuery = True		
End Function


'========================================================================================
Function FncPrint()
    Call Parent.FncPrint()
End Function


'========================================================================================
Function FncExcel()
	Call Parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
Function FncFind()
    Call Parent.FncFind(Parent.C_MULTI , False)
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

'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
	        
    With frm1
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtFromGlDt=" & Trim(.txtFromGlDt.Text)
		strVal = strVal & "&txtToGlDt=" & Trim(.txtToGlDt.Text)
		strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.Value)
		strVal = strVal & "&txtBizAreaCd1=" & Trim(.txtBizAreaCd1.Value)
		strVal = strVal & "&txtFrAcctCd=" & Trim(.txtFrAcctCd.Value)
		strVal = strVal & "&txtToAcctCd=" & Trim(.txtToAcctCd.Value)
		strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt1=" & Trim(.txtBizAreaCd1.Alt)
		strVal = strVal & "&txtFrAcctCd_Alt=" & Trim(.txtFrAcctCd.Alt)
		strVal = strVal & "&txtToAcctCd_Alt=" & Trim(.txtToAcctCd.Alt)
		
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------

		strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSqlGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
			' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동       
        	
    End With
    
    DbQuery = True


End Function


'========================================================================================
Function DbQueryOk()
'    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	IF Trim(frm1.txtBizAreaCd.value) = "" then
		frm1.txtBizAreaNm.value = ""
	end if
	
	IF Trim(frm1.txtBizAreaCd1.value) = "" then
		frm1.txtBizAreaNm1.value = ""
	end if	
	
	Call FncSetToolBar("Query")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function



'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolBar("1100000000001111")
	Case "QUERY"
		Call SetToolBar("1100000000011111")
	End Select
End Function

'========================================================================================
Sub  txtFrAcctCd_OnChange()
	IF Trim(frm1.txtFrAcctCd.value) = "" THEN
		frm1.txtFrAcctNm.value = ""
	ELSE
		IF Trim(frm1.txtToAcctCd.value) = "" THEN
			frm1.txtToAcctCd.value = frm1.txtFrAcctCd.value
			frm1.txtToAcctNm.value = frm1.txtFrAcctNm.value
		END IF	
	END IF		
End Sub


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>회계일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="시작회계일자" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
												           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToGlDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="종료회계일자" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP></TD>						
								    <TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP ><INPUT TYPE=TEXT NAME="txtFrAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="시작계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:CALL OpenPopUp(frm1.txtFrAcctCd.value, 1)">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtFrAcctNm" SIZE=25 tag="14">&nbsp;~&nbsp;														   
								    </TD>
								    <TD CLASS="TD5" NOWRAP>사업장코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizAreaCd" ALT="사업장코드" Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: LEFT" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenPopUp(frm1.txtBizAreaCd.value, 0)">
														   <INPUT NAME="txtBizAreaNm" ALT="사업장명" Size="25" MAXLENGTH="40" STYLE="TEXT-ALIGN: LEFT" tag="14NXXU">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP ><INPUT TYPE=TEXT NAME="txtToAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="종료계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:CALL OpenPopUp(frm1.txtToAcctCd.value, 2)">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtToAcctNm" SIZE=25 tag="14">																							   
								    </TD>														   				
								    <TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBizAreaCd1" ALT="사업장코드" Size="12" MAXLENGTH="10" STYLE="TEXT-ALIGN: LEFT" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenPopUp(frm1.txtBizAreaCd.value, 3)">
														   <INPUT NAME="txtBizAreaNm1" ALT="사업장명" Size="25" MAXLENGTH="40" STYLE="TEXT-ALIGN: LEFT" tag="14NXXU"></TD>					
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
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
													
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발생금액</TD>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TD5 NOWRAP>차이</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>누계금액</TD>
								<TD CLASS=TD5 NOWRAP>차변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TD5 NOWRAP>대변</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
													
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1">
</TEXTAREA><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
