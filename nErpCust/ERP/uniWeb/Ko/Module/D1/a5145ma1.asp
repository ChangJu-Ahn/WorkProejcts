<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5145MA1
'*  4. Program Name         : 유형별계정코드조회 
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
'	1. Constant는 반드시 대문자 표기.
'==========================================================================================
Dim lgIsOpenPop
Dim IsOpenPop                                               '☜: Popup status
Dim lgMark                                                  '☜: 마크 
Dim  gSelframeFlg
Dim lgPageNo2
Dim lgIntFlgMode2
'==========================================================================================
Const BIZ_PGM_ID		= "A5145MB1.asp"
Const BIZ_PGM_ID2		= "A5145MB2.asp"
'Dim lsPoNo								                       '☆: Jump시 Cookie로 보낼 Grid value
Const C_MaxKey          = 5                                    '☆☆☆☆: Max key value

Const TAB1 = 1													'☜: Tab의 위치 
Const TAB2 = 2

'==========================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
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
'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
End Sub

'==========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'==========================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("A5145MA1","S","A","V20021212",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    Call SetZAdoSpreadSheet("A5145MA102","S","B","V20021213",parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
    Call SetSpreadLock("B")
End Sub



'==========================================================================================
Sub SetSpreadLock(ByVal pOpt)
	With frm1
		If pOpt = "A" Then
			.vspdData.ReDraw = False
			ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
		ElseIf pOpt = "B" Then
			.vspdData2.ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.vspdData2.MaxCols,-1,-1
			.vspdData2.ReDraw = True
		End if
	End With
End Sub



'==========================================================================================
Sub InitComboBox()
	Err.clear
End Sub
 


'==========================================================================================
'	Name : OpenTransType()
'	Description : Plant PopUp
'==========================================================================================
Function OpenClassType(strCode, iwhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strVar
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정분류형태 팝업"
	arrParam(1) = "A_ACCT_CLASS_TYPE"
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "계정분류형태"

    arrField(0) = "CLASS_TYPE"
	arrField(1) = "CLASS_TYPE_NM"

    arrHeader(0) = "계정분류형태코드"
	arrHeader(1) = "계정분류형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iwhere
		Case 1
			frm1.txtClassType.focus
		Case 2
			frm1.txtClassType2.focus
		End Select
		Exit Function
	Else
		Call SetClass(arrRet,iwhere)
	end if

End Function
'==========================================================================================
'	Name : SetClass()
'	Description : Item Popup에서 Return되는 값 setting
'==========================================================================================
Function SetClass(Byval arrRet,Byval iwhere)
	With frm1
		Select Case iwhere
		Case 1
			.txtClassType.focus
			.txtClassType.value = Trim(arrRet(0))
			.txtClassTypeNm.value = arrRet(1)
		Case 2
			.txtClassType2.focus
			.txtClassType2.value = Trim(arrRet(0))
			.txtClassTypeNm2.value = arrRet(1)
		End Select
	End With
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

	arrParam(0) = "계정 팝업"										' 팝업 명칭 
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
'	arrField(4) = "HH" & parent.gColSep & "A_Acct.bal_fg"			' Field명(3)

	arrHeader(0) = "계정코드"										' Header명(0)
	arrHeader(1) = "계정코드명"									' Header명(1)
	arrHeader(2) = "그룹코드"										' Header명(2)
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
		case 4
			frm1.txtAcctCd21.focus
		End select

		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If

End Function
'==========================================================================================
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'==========================================================================================
Function SetReturnVal(ByVal arrRet,ByVal field_fg) 
	With frm1
		Select case field_fg
			case 1
				.txtBizAreaCd.focus
				.txtBizAreaCd.Value		= Trim(arrRet(0))
				.txtBizAreaNm.Value		= arrRet(1)
			case 2
				.txtAcctCd.focus
				.txtAcctCd.Value		= Trim(arrRet(0))
				.txtAcctNm.Value		= Trim(arrRet(1))
				.txtAcctCd2.Value		= arrRet(0)
				.txtAcctNm2.Value		= arrRet(1)

			case 3
				.txtAcctCd2.focus
				.txtAcctCd2.Value		= arrRet(0)
				.txtAcctNm2.Value		= arrRet(1)
			case 4
				.txtAcctCd21.focus
				.txtAcctCd21.Value		= Trim(arrRet(0))
				.txtAcctNm21.Value		= Trim(arrRet(1))
				.txtAcctCd22.Value		= Trim(arrRet(0))
				.txtAcctNm22.Value		= Trim(arrRet(1))
			case 5
				.txtAcctCd22.focus
				.txtAcctCd22.Value		= Trim(arrRet(0))
				.txtAcctNm22.Value		= Trim(arrRet(1))

		End select
	End With

End Function

'==========================================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	Dim gPos

	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       Case "VSPDDATA2"
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
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call FncSetToolBar("New")
	Call ClickTab1()
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'==========================================================================================
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



'==========================================================================================
' Tab 2
'==========================================================================================
Sub txtClassType2_onKeyPress()
    If window.event.keycode = 13 Then
        Call fncQuery()
    End If
End Sub

Sub txtAcctCd21_onKeyPress()
    If window.event.keycode = 13 Then
        Call fncQuery()
    End If
End Sub

Sub txtAcctCd22_onKeyPress()
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


'==========================================================================================
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
Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SP2C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData2

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
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
	Call SetSpreadColumnValue("B", frm1.vspdData2, Col, Row)
End Sub


Sub vspdData2_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then
    End If
End Sub

'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
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
Function FncQuery() 

    FncQuery = False
    Err.Clear

    Call InitVariables 

	If gSelframeFlg = TAB1 Then
		If Not chkField(Document, "1") Then
		   Exit Function
		End If
		Call ggoOper.ClearField(Document, "2")
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
	ElseIf gSelframeFlg = TAB2 Then
		If Not chkField(Document, "3") Then
		   Exit Function
		End If

		Call ggoOper.ClearField(Document, "4")
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	End If

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
    Call parent.FncFind(parent.C_MULTI , False)
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
Function DbQuery()
	Dim strVal, strZeroFg

    DbQuery = False

    Err.Clear
	Call LayerShowHide(1)
	        
    With frm1
    If gSelframeFlg = TAB1 Then 
		If lgIntFlgMode  <> Parent.OPMD_UMODE Then   ' This means that it is first search
			strVal = BIZ_PGM_ID & "?txtClassType=" & Trim(Frm1.txtClassType.value)
			If frm1.ProcessOpt1.checked = True Then
				strVal = strVal & "&ProcessOption=" & "1"	' All
			Else
				strVal = strVal & "&ProcessOption=" & "2"
			End If
		Else
			strVal = BIZ_PGM_ID & "?txtClassType=" & Trim(Frm1.htxtClassType.value)
			If frm1.hProcessOpt1.value = "1" Then
				strVal = strVal & "&ProcessOption=" & "1"	' All
			Else
				strVal = strVal & "&ProcessOption=" & "2"
			End If
        End If
		strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	ElseIf gSelframeFlg = TAB2 Then
		If lgIntFlgMode2  <> Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID2 & "?txtClassType=" & Trim(Frm1.txtClassType2.value)
			strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd21.Value)
			strVal = strVal & "&txtAcctCd2=" & Trim(.txtAcctCd22.Value)
			strVal = strVal & "&txtAcctCd_Alt=" & Trim(.txtAcctCd21.Alt)
		Else
			strVal = BIZ_PGM_ID2 & "?txtClassType=" & Trim(Frm1.htxtClassType2.value)
			strVal = strVal & "&txtAcctCd=" & Trim(.htxtAcctCd21.Value)
			strVal = strVal & "&txtAcctCd2=" & Trim(.htxtAcctCd22.Value)
			strVal = strVal & "&txtAcctCd_Alt=" & Trim(.txtAcctCd21.Alt)
        End If
		strVal = strVal & "&lgPageNo2="   & lgPageNo2                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
	End If
	Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function



'==========================================================================================
Function DbQueryOk()
	Call FncSetToolBar("Query")
    lgIntFlgMode     = Parent.OPMD_UMODE
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function

'==========================================================================================
Function DbQueryOk2()
    lgIntFlgMode2     = Parent.OPMD_UMODE
	Call FncSetToolBar("Query")
	frm1.vspdData2.focus
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

'==========================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab
	gSelframeFlg = TAB1
	frm1.txtClassType.focus
	'Call SetDefaultVal()
End Function


'==========================================================================================
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ 두번째 Tab
	gSelframeFlg = TAB2
	frm1.txtClassType2.focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>유형별계정코드조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>미등록계정코드</font></td>
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
		<DIV ID="TabDiv"  SCROLL="no">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">계정분류형태</TD>
									<TD CLASS="TD6">
										<INPUT NAME="txtClassType" MAXLENGTH="4" SIZE=12 ALT ="계정분류형태" tag="13XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:call OpenClassType(frm1.txtClassType.value,1)">
										<INPUT NAME="txtClassTypeNm" MAXLENGTH="50" SIZE=20 ALT ="" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>조회내역</TD>
									<TD CLASS="TD6" NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ProcessOpt" CHECKED ID="ProcessOpt1" VALUE="Y" tag="12"><LABEL FOR="ProcessOpt1">전체</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ProcessOpt" ID="ProcessOpt2" VALUE="N" tag="12"><LABEL FOR="ProcessOpt2">중복계정코드</LABEL></SPAN>
									</TD>
								</TR>
<!--
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP> <INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd.value,2)"> <INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="24">&nbsp;</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>

								 </TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP> <INPUT TYPE=TEXT NAME="txtAcctCd2" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd2.value,3)"> <INPUT TYPE=TEXT NAME="txtAcctNm2" SIZE=25 tag="24"></TD>
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
								<TD HEIGHT="100%" colspan=7>
								<script language =javascript src='./js/a5145ma1_vspdData_vspdData.js'></script></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</div>
		<!--두번째 TAB  -->
		<DIV ID="TabDiv"  SCROLL="no">

			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">계정분류형태</TD>
									<TD CLASS="TD6">
										<INPUT NAME="txtClassType2" MAXLENGTH="4" SIZE=12 ALT ="계정분류형태" tag="33XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:call OpenClassType(frm1.txtClassType2.value,2)">
										<INPUT NAME="txtClassTypeNm2" MAXLENGTH="50" SIZE=20 ALT ="" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtAcctCd21" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="31XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd21.value,4)"> <INPUT TYPE=TEXT NAME="txtAcctNm21" SIZE=25 tag="24">&nbsp; ~</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>

								 </TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd22" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="31XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtAcctCd22.value,5)"> <INPUT TYPE=TEXT NAME="txtAcctNm22" SIZE=25 tag="24"></TD>
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
								<TD HEIGHT="100%" colspan=7>
								<script language =javascript src='./js/a5145ma1_vspdData2_vspdData2.js'></script></TD>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtClassType"		tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="hProcessOpt1"		tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="htxtClassType2"	tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="htxtAcctCd21"		tag="24" TABINDEX="-1" >
<INPUT TYPE=HIDDEN NAME="htxtAcctCd22"		tag="24" TABINDEX="-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
 

