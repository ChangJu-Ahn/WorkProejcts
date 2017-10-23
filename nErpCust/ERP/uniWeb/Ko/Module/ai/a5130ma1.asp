<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Closing and Financial Statements
'*  3. Program ID           : a5130ma1
'*  4. Program Name         : 기초일계표조회 
'*  5. Program Desc         : Query of Daily/Monthly Summerization
'*  6. Component List       :
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Soo Min, Oh
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">				  </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

'========================================================================================
Dim lgStrPrevKey                                            '☜: Next Key tag                          
Dim lgSortKey                                               '☜: Sort상태 저장변수                      
Dim IsOpenPop                                               '☜: Popup status                           


Dim lgMark                                                  '☜: 마크 
Dim strDateYr
Dim strDateMonth
Dim strDateDay                                  

Dim lgIsOpenPop 

'========================================================================================
Const BIZ_PGM_ID        = "A5130MB1.asp"
Const C_MaxKey          = 1 

'========================================================================================
Sub InitVariables()
    lgStrPrevKey     = ""
    lgSortKey        = 1

End Sub

'========================================================================================
Sub SetDefaultVal()
	frm1.fpDateYr.Text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)	
	Call ggoOper.FormatDate(frm1.txtDateYr,  parent.gDateFormat, 3)
	frm1.txtDateYr.focus
End Sub

'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "COOKIE", "QA") %>
End Sub


'========================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("A5130MA1", "S", "A", "V20030102", parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
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

	Select Case iWhere
	Case 0
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
    
	Case 1
		arrParam(0) = "일계표유형 팝업"					' 팝업 명칭 
		arrParam(1) = "A_ACCT_CLASS_TYPE"						' TABLE 명칭 
		arrParam(2) = strCode									' Code Condition
		arrParam(3) = ""										' Name Cindition
		arrParam(4) = "CLASS_TYPE LIKE " & FilterVar("DMS%", "''", "S") & " "										' Where Condition
		arrParam(5) = "일계표유형"			
	
	    arrField(0) = "CLASS_TYPE"								' Field명(0)
		arrField(1) = "CLASS_TYPE_NM"							' Field명(1)
    
	    arrHeader(0) = "일계표유형"						' Header명(0)
		arrHeader(1) = "일계표유형명"							' Header명(1)
    
	Case Else
		Exit Function
	End Select
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		Case 0
			frm1.txtBizAreaCd.focus
		Case 1
			frm1.txtClassType.focus
		End Select
		Exit Function
	Else
		Call SetReturnVal(arrRet, iWhere)
	End If	

End Function

'========================================================================================
Function SetReturnVal(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
		Case 0
			.txtBizAreaCd.focus
			.txtBizAreaCd.value = arrRet(0)
			.txtBizAreaNm.value = arrRet(1)
		Case 1
			.txtClassType.focus
			.txtClassType.value   = arrRet(0)
			.txtClassTypeNm.value = arrRet(1)
		End Select
	End With

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
Function OpenPopupAcct()

	Dim arrRet
	Dim arrParam(5)	
	Dim IntRetCD
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5130ra2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra2", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function


	arrParam(0) = Trim(GetKeyPosVal("A", 1))	'계정코드 
	arrParam(1) = frm1.txtDateYr.Text	'년도 
	arrParam(2) = frm1.txtDateMnth.value '월 
	arrParam(3) = frm1.txtBizAreaCd.value '사업장 
	IsOpenPop = True
    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
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
	Call FncSetToolBar("New")
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================
Sub txtDateYr_DblClick(Button)
	if Button = 1 then
		frm1.txtDateYr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateYr.Focus
	End if
End Sub

'========================================================================================
Sub txtDateYr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtBizAreaCd.focus
	   Call FncQuery()
	End If
End Sub

'========================================================================================
Sub txtDateYr_Change()
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

	If Row < 1 Then Exit Sub
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
    	If lgStrPrevKey <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           
           If DbQuery2 = False Then
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

'========================================================================================
Function FncQuery()

    FncQuery = False
    Err.Clear

    Call ggoOper.ClearField(Document, "2")
    Call InitVariables

    If Not chkField(Document, "1") Then
       Exit Function
    End If
		
	Call ExtractDateFrom(frm1.txtDateYr.Text,frm1.txtDateYr.UserDefinedFormat,parent.gComDateType,strDateYr,strDateMonth,strDateDay)

    IF  DbQuery	= False Then
		Exit Function
	END IF
	
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


'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False

    Err.Clear
	Call LayerShowHide(1)
	Call FncSetToolBar("Query")

    With frm1
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtDateYr=" & strDateYr
		strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.Value)
		strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtDateMnth=" & Trim(.txtDateMnth.value)

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------

		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)
    End With    
    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()
	IF Trim(frm1.txtBizAreaCd.value) = "" then
		frm1.txtBizAreaNm.value = ""
	end if
	Call FncSetToolBar("NEW")
	CALL vspdData_Click(1, 1)
	frm1.vspdData.focus
	Set gActiveElement = document.activeElement 
End Function

'========================================================================================
Sub SetPrintCond(StrEbrFile, VarBizArea, VarClassTypeFr, VarClassTypeTo, VarDateFr, VarDateTo, VarBalTAmt)
	Dim strGlYear
	Dim strGlMonth
	Dim strgGlDay
	Dim strFiscYr,strFiscMnth,strFiscDt

	StrEbrFile = "a5130ma1"

	With frm1

		If Trim(.txtBizAreaCd.value) = "" Then
			VarBizArea = "*"
		Else
			VarBizArea = UCase(Trim(.txtBizAreaCd.value))
		End If

		Call ExtractDateFrom(frm1.txtDateYr.Text,frm1.txtDateYr.UserDefinedFormat,parent.gComDateType,strGlYear,strGlMonth,strgGlDay)
		Call ExtractDateFrom(parent.gFiscStart,parent.gAPDateFormat,parent.gComDateType,strFiscYr,strFiscMnth,strFiscDt)
		
		VarClassTypeFr = UCase(Trim(.txtClassType.value))
		VarDateFr = strGlYear & strFiscMnth & strFiscDt
		VarDateTo = UNIDateAdd("D" ,+364, VarDateFr, parent.gServerDateFormat)
		VarBalTAmt = Replace(.txtTAmt.value, parent.gComNum1000, "")
	End With
	
End Sub

'========================================================================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile, VarBizArea, VarClassTypeFr, VarClassTypeTo, VarDateFr, VarDateTo, VarBalTAmt
	Dim ObjName
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

'	If UniConvDateToYYYYMMDD(frm1.txtDateYr.Text, parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, parent.gDateFormat, "") Then
'		Call DisplayMsgBox("970025", "X", frm1.txtDateYr.Alt, frm1.txtDateTo.Alt)
'		frm1.txtDateYr.focus
'		Exit Function
'	End If
	
	Call SetPrintCond(StrEbrFile, VarBizArea, VarClassTypeFr, VarClassTypeTo, VarDateFr, VarDateTo, VarBalTAmt)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
'    On Error Resume Next                                                    '☜: Protect system from crashing
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "BizArea|" & VarBizArea
	StrUrl = StrUrl & "|ClassTypeFr|" & VarClassTypeFr
'	StrUrl = StrUrl & "|ClassTypeTo|" & VarClassTypeTo
	StrUrl = StrUrl & "|DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BalTAmt|" & VarBalTAmt

	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function


'========================================================================================
Function FncBtnPreview() 
	'On Error Resume Next

	Dim StrUrl
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile, VarBizArea, VarClassTypeFr, VarClassTypeTo, VarDateFr,VarDateTo,  VarBalTAmt
	Dim ObjName
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
	Call SetPrintCond(StrEbrFile, VarBizArea, VarClassTypeFr, VarClassTypeTo, VarDateFr, VarDateTo, VarBalTAmt)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
	StrUrl = StrUrl & "BizArea|" & VarBizArea
	StrUrl = StrUrl & "|ClassTypeFr|" & VarClassTypeFr
	StrUrl = StrUrl & "|DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BalTAmt|" & VarBalTAmt

	Call FncEBRPreview(ObjName,StrUrl)
		
End Function
'========================================================================================
'툴바버튼 세팅 
'========================================================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1000000000011111")
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
								<td background="../../image/table/seltab_up_bg.gif" NOWRAP><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기초일계표조회</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right><A HREF="VBSCRIPT:OpenPopupAcct()">계정별보조부조회</A>&nbsp;</td>
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
									<TD CLASS="TD5" NOWRAP>회계년도</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5130ma1_fpDateYr_txtDateYr.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>사업장코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtBizAreaCd.value,0)">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="24X" ALT="사업장명" STYLE="TEXT-ALIGN: Left">
									</TD>
								</TR>
                                <TR>
									<TD NOWRAP CLASS="TD5">월</TD>
									<TD NOWRAP CLASS="TD6"><SELECT NAME="txtDateMnth" TAG="11XXXU" ALT="월">
									                        <OPTION VALUE=""></OPTION>
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
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/a5130ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>차변합계</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDrTAmt" TYPE="Text" MAXLENGTH="20" STYLE="TEXT-ALIGN: right" tag="24X2"></TD>
								<TD CLASS=TD5 NOWRAP>대변합계</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCrTAmt" TYPE="Text" MAXLENGTH="20" STYLE="TEXT-ALIGN: right" tag="24X2"></TD>
							</TR>
<!--
							<TR>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
								<TD CLASS=TD5 NOWRAP>현금잔액</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTAmt" TYPE="Text" MAXLENGTH="20" STYLE="TEXT-ALIGN: right" tag="24X2"></TD>
							</TR>
-->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
<!--
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
-->
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBizAreaCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hClassType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hClassCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDateFr" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDateTo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCommand" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>
