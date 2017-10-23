<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S7111QA1
'*  4. Program Name         : NEGO 현황조회 
'*  5. Program Desc         : NEGO 현황조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2002/12/14 Include 성능향상 강준구 
'********************************************************************************************** 
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 
Dim IscookieSplit 

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s7111qb1.asp"
Const BIZ_PGM_JUMP_ID	= "s7111ma1"
Const C_MaxKey          = 3                                    '☆☆☆☆: Max key value
                                            '☆: Jump시 Cookie로 보낼 Grid value
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
<%'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------%>
	frm1.txtNegoFrDt.text = StartDate
	frm1.txtNegoToDt.text = EndDate

<%'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------%>
	frm1.txtSalesGroup.focus 
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S7111QA1","S","A","V20030321", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'========================================================================================================= 
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================= 
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				<%' Where Condition%>
		arrParam(5) = "수입자"							<%' TextBox 명칭 %>
	
		arrField(0) = "BP_CD"								<%' Field명(0)%>
		arrField(1) = "BP_NM"								<%' Field명(1)%>
    
		arrHeader(0) = "수입자"							<%' Header명(0)%>
		arrHeader(1) = "수입자명"						<%' Header명(1)%>

		frm1.txtconBp_cd.focus
	Case 1
		arrParam(1) = "B_MINOR"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtPayTerms.Value)			<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD = " & FilterVar("B9004", "''", "S") & ""					<%' Where Condition%>
		arrParam(5) = "결제방법"							<%' TextBox 명칭 %>
	
		arrField(0) = "MINOR_CD"								<%' Field명(0)%>
		arrField(1) = "MINOR_NM"								<%' Field명(1)%>
    
		arrHeader(0) = "결제방법"							<%' Header명(0)%>
		arrHeader(1) = "결제방법명"							<%' Header명(1)%>

		frm1.txtPayTerms.focus
	Case 3
		arrParam(1) = "B_SALES_GRP"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtSalesGroup.Value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "영업그룹"						<%' TextBox 명칭 %>
	
		arrField(0) = "SALES_GRP"							<%' Field명(0)%>
		arrField(1) = "SALES_GRP_NM"							<%' Field명(1)%>
    
		arrHeader(0) = "영업그룹"						<%' Header명(0)%>
		arrHeader(1) = "영업그룹명"							<%' Header명(1)%>

		frm1.txtSalesGroup.focus
	End Select

	arrParam(0) = arrParam(5)								<%' 팝업 명칭 %>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function
'========================================================================================================= 
Function PopZAdoConfigGrid()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
'========================================================================================================= 
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_Nm.value = arrRet(1)   
		Case 1
			.txtPayTerms.value = arrRet(0) 
			.txtPayTermsNm.value = arrRet(1)   
		Case 2
			.txtSalesOrg.value = arrRet(0)
			.txtSalesOrgNm.value = arrRet(1)  
		Case 3
			.txtSalesGroup.value = arrRet(0) 
			.txtSalesGroupNm.value = arrRet(1)   
		End Select
	End With
End Function
'========================================================================================================= 
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						<% 'Cookie Split String : CookiePage Function Use%>

	If Kubun = 1 Then								<% 'Jump로 화면을 이동할 경우 %>

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		
		WriteCookie CookieSplit , IsCookieSplit					<% 'Jump로 화면을 이동할때 필요한 Cookie 변수정의 %>
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							<% 'Jump로 화면이 이동해 왔을경우 %>

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If
		
		Dim iniSep

<%'--------------- 개발자 coding part(실행로직,Start)---------------------------------------------------%>
		<% '자동조회되는 조건값과 검색조건부 Name의 Match %>
		frm1.txtconBp_cd.value =  arrVal(0)
		frm1.txtconBp_Nm.value =  arrVal(1)
		frm1.txtBillType.value =  arrVal(2)
		frm1.txtBillTypeNm.value = arrVal(3) 
		frm1.txtSalesOrg.value =  arrVal(4)
		frm1.txtSalesOrgNm.value = arrVal(5) 
		frm1.txtSalesGroup.value =  arrVal(6)
		frm1.txtSalesGroupNm.value = arrVal(7) 
		frm1.txtItem_cd.value =  arrVal(8)
		frm1.txtItem_Nm.value = arrVal(9)

<%'--------------- 개발자 coding part(실행로직,End)---------------------------------------------------%>

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call FncQuery()

		WriteCookie CookieSplit , ""

	End IF
End Function

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
End Sub
'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort In Assending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort In Desending
			lgSortKey = 1
		End If
		Exit Sub
	End If

<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------%>
	If Row < 1 Then Exit Sub
	
	IscookieSplit = ""
	
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 1
    IscookieSplit = frm1.vspdData.text
	
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------%>
    
End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then  Exit Sub
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
   End if
    
End Sub
'========================================================================================================= 
	Sub rdoFlawExistFlg1_OnClick()
		frm1.txtRadio.value = frm1.rdoFlawExistFlg1.value
	End Sub

	Sub rdoFlawExistFlg2_OnClick()
		frm1.txtRadio.value = frm1.rdoFlawExistFlg2.value
	End Sub

	Sub rdoFlawExistFlg3_OnClick()
		frm1.txtRadio.value = frm1.rdoFlawExistFlg3.value
	End Sub
	
'========================================================================================================= 
	Sub txtNegoFrDt_DblClick(Button)
		If Button = 1 Then
			frm1.txtNegoFrDt.Action = 7
			Call SetFocusToDocument("M")	
			Frm1.txtNegoFrDt.Focus
		End If
	End Sub

	Sub txtNegoToDt_DblClick(Button)
		If Button = 1 Then
			frm1.txtNegoToDt.Action = 7
			Call SetFocusToDocument("M")	
			Frm1.txtNegoToDt.Focus
		End If
	End Sub
'========================================================================================================= 
	Sub txtNegoFrDt_KeyDown(KeyCode, Shift)
		If KeyCode = 13	Then Call FncQuery()
	End Sub
	Sub txtNegoToDt_KeyDown(KeyCode, Shift)
		If KeyCode = 13	Then Call FncQuery()
	End Sub

'========================================================================================================= 
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    lgIntFlgMode = parent.OPMD_CMODE	
	If ValidDateCheck(frm1.txtNegoFrDt, frm1.txtNegoToDt) = False Then Exit Function
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    

    '-----------------------
    'Query function call area
    '-----------------------

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'========================================================================================================= 
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================================= 
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================================= 
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================================= 
Sub FncSplitColumn()    
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================================= 
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================================= 
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1

<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = strVal & "&txtconBp_cd=" & Trim(.txtHconBp_cd.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtHSalesGroup.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.txtHPayTerms.value)
			strVal = strVal & "&txtNegoFrDt=" & Trim(.txtHNegoFrDt.value)
			strVal = strVal & "&txtNegoToDt=" & Trim(.txtHNegoToDt.value)
			strVal = strVal & "&txtRadio=" & Trim(.txtHRadio.value)
		Else
			strVal = strVal & "&txtconBp_cd=" & Trim(.txtconBp_cd.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.txtPayTerms.value)
			strVal = strVal & "&txtNegoFrDt=" & Trim(.txtNegoFrDt.text)
			strVal = strVal & "&txtNegoToDt=" & Trim(.txtNegoToDt.text)
			strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)		
		End If
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
       
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True


End Function

'========================================================================================================= 
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
'    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call SetToolbar("11000000000111")
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			Call vspdData_Click(1, 1)
		End If
		lgIntFlgMode = parent.OPMD_UMODE			
    Else
       frm1.txtSalesGroup.focus
    End If     

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Nego현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*">&nbsp;</td>
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
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
									<TD CLASS="TD5" NOWRAP>수입자</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="수입자" TYPE="Text" MAXLENGTH=10 SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>결제방법</TD>
									<TD CLASS=TD6><INPUT NAME="txtPayTerms" ALT="결제방법" TYPE="Text" MAXLENGTH=5 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT NAME="txtPaytermsNm" TYPE="Text" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>NEGO일</TD>
									<TD CLASS=TD6 NOWRAP>
									
										<script language =javascript src='./js/s7111qa1_fpDateTime1_txtNegoFrDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s7111qa1_fpDateTime2_txtNegoToDt.js'></script>
									</TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>입금유무</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlawExistFlg" TAG="11X" VALUE="A" CHECKED ID="rdoFlawExistFlg1"><LABEL FOR="rdoFlawExistFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlawExistFlg" TAG="11X" VALUE="Y" ID="rdoFlawExistFlg2"><LABEL FOR="rdoFlawExistFlg2">Y</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFlawExistFlg" TAG="11X" VALUE="N" ID="rdoFlawExistFlg3"><LABEL FOR="rdoFlawExistFlg3">N</LABEL>			
									</TD>
									<TD CLASS=TD5 NOWRAP></TD> 
									<TD CLASS=TD6 NOWRAP></TD>
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
								<TD HEIGHT="100%">
								<script language =javascript src='./js/s7111qa1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<td <%=HEIGHT_TYPE_01%>></td>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">NEGO등록</a></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="HconItem_cd" tag="24" TABINDEX="-1"> 
<INPUT TYPE=HIDDEN NAME="HValid_from_dt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconCurrency" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconDeal_type" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconPay_terms" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HconSales_unit" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHconBp_cd" tag="24" TABINDEX="-1"> 
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHNegoFrDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHNegoToDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHRadio" tag="24" TABINDEX="-1">


</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
