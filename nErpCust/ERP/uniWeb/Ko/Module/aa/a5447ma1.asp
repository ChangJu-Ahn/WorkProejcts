<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Accounting
*  2. Function Name        : 
*  3. Program ID           : A5447MA1
*  4. Program Name         : 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2003/05/28
*  8. Modified date(Last)  : 2003/05/28
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<Script Language="VBScript">

Option Explicit                                                        '☜: Turn on the Option Explicit option.


'========================================================================================================

Const BIZ_PGM_ID 		= "A5447MB1.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================


'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
Dim lgIsOpenPop                                          
Dim IsOpenPop  
Dim lgCookValue
Dim lgStrColorFlag
Dim lgSaveRow 

Const C_MaxKey          = 2


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
	Dim strSvrDate
	frm1.txtFrDt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.txtFrDt, parent.gDateFormat, 2)
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>

End Sub


'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면의 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim strTemp, arrVal

	Const CookieSplit = 4877

	If Kubun = 0 Then                                              ' Called Area
      strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, parent.gRowSep)

       Frm1.txtSchoolCd.Value = ReadCookie ("SchoolCd")
       Frm1.txtGrade.Value   = arrVal(0)
       Call MainQuery()

       WriteCookie CookieSplit , ""

	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(Frm1.vspdData.ActiveCol,Frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function


'========================================================================================================
Sub InitSpreadSheet()
		Call SetZAdoSpreadSheet("A5447MA1","S","A","V20030631",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
		Call SetSpreadLock()
End Sub


'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True
	End With
End Sub



'========================================================================================================
Sub InitComboBox()
End Sub

'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
End Sub

'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000011111")										
	Call InitComboBox
'    Call CookiePage(0)
    Frm1.txtFrDt.focus
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
Function FncQuery() 

    On Error Resume Next
    Err.Clear

    FncQuery = False

    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()

    Call InitVariables

    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If DbQuery = False Then
       Exit Function
    End If

    If Err.number = 0 Then
       FncQuery = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncNew()
End Function
	
'========================================================================================================
Function FncDelete()
End Function


'========================================================================================================
Function FncSave()
End Function

'========================================================================================================
Function FncCopy()
End Function

'========================================================================================================
Function FncCancel() 
End Function

'========================================================================================================
Function FncInsertRow()
End Function

'========================================================================================================
Function FncDeleteRow()
End Function

'========================================================================================================
Function FncPrint()

    On Error Resume Next
    Err.Clear

    FncPrint = False
	Call Parent.FncPrint()

    If Err.number = 0 Then
       FncPrint = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncPrev()
End Function

'========================================================================================================
Function FncNext()
End Function

'========================================================================================================
Function FncExcel()

    On Error Resume Next
    Err.Clear

    FncExcel = False

	Call Parent.FncExport(parent.C_MULTI)

    If Err.number = 0 Then
       FncExcel = True
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncFind() 
    On Error Resume Next
    Err.Clear
    FncFind = False
	Call Parent.FncFind(parent.C_MULTI, True)
    If Err.number = 0 Then
       FncFind = True
    End If
    Set gActiveElement = document.ActiveElement 
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
    On Error Resume Next
    Err.Clear
    FncExit = False
	If Err.number = 0 Then
       FncExit = True
    End If
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function DbQuery() 
	Dim strVal
	Dim strYear,strMonth,strDay
	Dim strYYYYMM
	On Error Resume Next
	Err.Clear

	DbQuery = False
	Call LayerShowHide(1)

	If lgIntFlgMode  <> parent.OPMD_UMODE Then
		Call ExtractDateFrom(frm1.txtFrDt.Text,frm1.txtFrDt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
		strYYYYMM = strYear & strMonth
	End If
	With frm1

		strVal = BIZ_PGM_ID	& "?txtMode="        & Parent.UID_M0001                      '☜: Query
		strVal = strVal		& "&txtMaxRows=" 	 & Frm1.vspdData.MaxRows				'☜: Max fetched data

	'--------- Developer Coding Part (Start) ----------------------------------------------------------
		If lgIntFlgMode  <> parent.OPMD_UMODE Then   ' This means that it is first search
		   strVal = strVal & "&txtFrDt="			& strYYYYMM
		   strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
		   strVal = strVal & "&txtFrAcctCd="		& Trim(.txtFrAcctCd.value)
		Else
		   strVal = strVal & "&txtFrDt="			& Trim(.htxtFrDt.value)
		   strVal = strVal & "&txtBizAreaCd="		& Trim(.htxtBizAreaCd.value)
		   strVal = strVal & "&txtFrAcctCd="		& Trim(.htxtFrAcctCd.value)
		End If

		strVal = strVal & "&lgPageNo="   & lgPageNo
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	'--------- Developer Coding Part (End) ------------------------------------------------------------

		Call RunMyBizASP(MyBizASP, strVal)

	End With

	If Err.number = 0 Then
	   DbQuery = True
	End If
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function DbQueryOk()												

    On Error Resume Next
    Err.Clear

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE
    lgSaveRow        = 1
	frm1.vspdData.focus

    Set gActiveElement = document.ActiveElement

	Call SetQuerySpreadColor

End Function

'========================================================================================================
'	Description : 스프레트시트의 특정 컬럼의 배경색상을 변경 
'========================================================================================================
Sub SetQuerySpreadColor()' Not nused

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	Dim Spread
	
	Set Spread = frm1.vspdData
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)
	
	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		Spread.Col = -1
		Spread.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				Spread.BackColor = RGB(204,255,153) '연두 
			Case "2"
				Spread.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				Spread.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				Spread.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				Spread.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next

End Sub


'========================================================================================================
	Function OpenBizAreaPopUp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장팝업"							' Popup Name
	arrParam(1) = "B_BIZ_AREA"								' Table Name
	arrParam(2) = Frm1.txtBizAreaCd.value					' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "사업장코드"

	arrField(0) = "BIZ_AREA_CD"								' Field명(0)
	arrField(1) = "BIZ_AREA_NM"								' Field명(1)

	arrHeader(0) = "사업장코드"							' Header명(0)
	arrHeader(1) = "사업장명"								' Header명(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtBizAreaCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizArea(arrRet)
	End If
End Function

'========================================================================================================
Sub SetBizArea(Byval arrRet)
	With Frm1
	  .txtBizAreaCd.value = Trim(arrRet(0))
	  .txtBizAreaNm.value = arrRet(1)
	End With
End Sub

'========================================================================================================
Function OpenAcctCd(byval strText, byval iwhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "자산계정팝업"							' Popup Name
	arrParam(1) = "A_ASSET_ACCT A, A_ACCT B"				' Table Name
	arrParam(2) = Trim(strText)								' Code Condition
	arrParam(3) = ""										' Name Cindition
	arrParam(4) = "A.ACCT_CD = B.ACCT_CD"					' Where Condition
	arrParam(5) = "자산계정"

	arrField(0) = "A.ACCT_CD"								' Field명(0)
	arrField(1) = "B.ACCT_NM"								' Field명(1)

	arrHeader(0) = "자산계정"								' Header명(0)
	arrHeader(1) = "자산계정명"							' Header명(1)

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If Cint(iwhere) = 1 then
		frm1.txtFrAcctCd.focus
	else
		frm1.txtToAcctCd.focus
	end if
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAcctCd(arrRet, iwhere)
	End If
End Function

'========================================================================================================
Sub SetAcctCd(ByVal arrRet, ByVal iwhere)
	With Frm1
		If Cint(iwhere) = 1 then
			.txtFrAcctCd.value = Trim(arrRet(0))
			.txtFrAcctNm.value = arrRet(1)
		else
			.txtToAcctCd.value = Trim(arrRet(0))
			.txtToAcctNm.value = arrRet(1)
			call txtToAcctCd_onChange()
		end if
	End With
End Sub





'==================================================================================
' Name : PopZAdoConfigGrid
' Desc :
'==================================================================================
Sub PopZAdoConfigGrid()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenOrderBy("A")

End Sub



'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet

	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()
   End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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
Sub vspdData_DblClick(ByVal Col, ByVal Row)
End Sub


'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub


'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub

'========================================================================================================
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtFrDt.Action = 7
       Call SetFocusToDocument("M")
       Frm1.txtFrDt.Focus
	End If
End Sub


'========================================================================================================
'   Event Name : fpdtFromEnterDt_KeyPress()
'   Event Desc : 
'========================================================================================================
Sub txtFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub




'========================================================================================================
'   Event Name : txtBizAreaCd_onChange
'   Event Desc : 
'========================================================================================================
Sub txtBizAreaCd_onChange()
'	Dim IntRetCD
'	Dim arrVal
'
'	If frm1.txtBizAreaCd.value = "" Then Exit Sub
'
'	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD= '" & TRim(frm1.txtBizAreaCd.value) & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
'		arrVal = Split(lgF0, Chr(11)) 
'		frm1.txtBizAreaNm.value= Trim(arrVal(0)) 
'	Else
'		IntRetCD = DisplayMsgBox("124200","X","X","X")
'		frm1.txtBizAreaCd.focus
'	End If
End Sub

'========================================================================================================
'   Event Name : txtAcctCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtFrAcctCd_onChange()
'	Dim IntRetCD
'	Dim arrVal
'
'	If CommonQueryRs("ACCT_NM", "A_ACCT ", " ACCT_CD= '" & TRim(frm1.txtFrAcctCd.value) & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
'		arrVal = Split(lgF0, Chr(11)) 
'		frm1.txtFrAcctNm.value= Trim(arrVal(0))
'	Else
'		IntRetCD = DisplayMsgBox("117100","X","X","X")
'		frm1.txtFrAcctCd.value = ""
'		frm1.txtFrAcctNm.value = ""
'		frm1.txtFrAcctCd.focus
'	End If
End Sub

'========================================================================================================
'   Event Name : txtAcctCd_onChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub txtFrAsstCd_onChange()
	Dim IntRetCD
	Dim arrVal
	
'	If CommonQueryRs("ASST_NM", "A_ASSET_MASTER ", " ASST_NO= '" & TRim(frm1.txtFrAsstCd.value) & "'", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
'		arrVal = Split(lgF0, Chr(11)) 
'		frm1.txtFrAsstNm.value= Trim(arrVal(0)) 
'	Else
'		IntRetCD = DisplayMsgBox("117400","X","X","X")
'		frm1.txtFrAsstCd.value = ""
'		frm1.txtFrAsstNm.value = ""
'		frm1.txtFrAsstCd.focus
'	End If
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고정자산계정별합계</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
							<!--
								<TR>
									<TD CLASS="TD5" NOWRAP>조회구분</TD>
									<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked onclick=radio1_onchange()><LABEL FOR=Rb_WK1>전표일자</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2 onclick=radio2_onchange()><LABEL FOR=Rb_WK2>거래일자</LABEL></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								-->
								<TR>
									<TD CLASS="TD5" ID = "TitleDate">년월</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5447ma1_fpDateTime1_txtFrDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaPopUp()"> <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=18 tag="14X" ALT="사업장명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>자산계정</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtFrAcctCd" SIZE=12 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="자산계정코드(From)"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFrAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenAcctCd(frm1.txtFrAcctCd.value, 1)"> <INPUT TYPE=TEXT NAME="txtFrAcctNm" SIZE=25 tag="14">&nbsp;</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5447ma1_vspdData_vspdData.js'></script>
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
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtFrDt"		tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"	tag="34" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtFrAcctCd"	tag="34" TABINDEX = "-1">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
