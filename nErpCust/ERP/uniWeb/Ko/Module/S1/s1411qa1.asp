<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1411QA1
'*  4. Program Name         : 여신현황조회 
'*  5. Program Desc         : 여신현황조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : -2002/12/12 : UI성능향상(include) 반영 강준구 
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

Dim prDBSYSDate
Dim EndDate ,StartDate

prDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = UniDateAdd("m", -1, EndDate,parent.gDateFormat)

Const BIZ_PGM_ID        = "s1411qb1.asp"
Const BIZ_PGM_JUMP_ID   = "s1411ma1"				  	       '☆: 비지니스 로직 ASP명 
Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value

Dim lsCreditGrp                                                 '☆: Jump시 Cookie로 보낼 Grid value

'========================================================================================================= 
Sub InitVariables()

    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
<%'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------%>
	frm1.txtCreditGrp.focus	
<%'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------%>

End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S1411QA1","S","A","V20021106", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'========================================================================================================= 
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

'========================================================================================================= 
Function OpenCreditGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "여신관리그룹"					<%' 팝업 명칭 %>
	arrParam(1) = "S_CREDIT_LIMIT"						<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtCreditGrp.value)			<%' Code Condition%>
	arrParam(3) = ""		                            <%' Name Cindition%>
	arrParam(4) = ""									<%' Where Condition%>
	arrParam(5) = "여신관리그룹"					<%' TextBox 명칭 %>
	
    arrField(0) = "CREDIT_GRP"							<%' Field명(0)%>
    arrField(1) = "CREDIT_GRP_NM"						<%' Field명(1)%>
    
    arrHeader(0) = "여신관리그룹"					<%' Header명(0)%>
    arrHeader(1) = "여신관리그룹명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
    frm1.txtCreditGrp.focus 

	If arrRet(0) = "" Then
		Exit Function
	Else 
		Call SetCreditGrp(arrRet)
	End If	

End Function

'========================================================================================================= 
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "화폐"							<%' 팝업 명칭 %>
	arrParam(1) = "B_CURRENCY"							<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtCurrency.Value)			<%' Code Condition%>
	arrParam(3) = ""		                            <%' Name Cindition%>
	arrParam(4) = ""									<%' Where Condition%>
	arrParam(5) = "화폐"							<%' TextBox 명칭 %>

	arrField(0) = "Currency"							<%' Field명(0)%>
	arrField(1) = "Currency_desc"						<%' Field명(1)%>

	arrHeader(0) = "화폐"							<%' Header명(0)%>
	arrHeader(1) = "화폐명"							<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
		
	frm1.txtCurrency.focus 

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurrency(arrRet)
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
Function SetCreditGrp(Byval arrRet)
	frm1.txtCreditGrp.value = arrRet(0)
	frm1.txtCreditGrpNm.value = arrRet(1)
End Function

'========================================================================================================= 
Function SetCurrency(arrRet)
	frm1.txtCurrency.Value = arrRet(0)
End Function

<% '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'==================================================================================================== %>
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						<% 'Cookie Split String : CookiePage Function Use%>

	If Kubun = 1 Then								<% 'Jump로 화면을 이동할 경우 %>

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie CookieSplit , lsCreditGrp					<% 'Jump로 화면을 이동할때 필요한 Cookie 변수정의 %>
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							<% 'Jump로 화면이 이동해 왔을경우 %>

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" Then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If
				
		Dim iniSep

<%'--------------- 개발자 coding part(실행로직,Start)---------------------------------------------------%>
		<% '자동조회되는 조건값과 검색조건부 Name의 Match %>
		For iniSep = 0 To UBound(arrVal) -1
			Select Case UCase(Trim(arrVal(iniSep)))
			Case UCase("여신관리그룹")
				frm1.txtCreditGrp.value =  arrVal(iniSep + 1)
			Case UCase("여신관리그룹명")
				frm1.txtCreditGrpNm.value =  arrVal(iniSep + 1)
			End Select
		Next
<%'--------------- 개발자 coding part(실행로직,End)---------------------------------------------------%>

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")							'⊙: 버튼 툴바 제어 
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------%>

	Call CookiePage(0)
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------%>
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

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",1) ' 1
	lsCreditGrp=frm1.vspdData.Text
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------%>
    
End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub


'========================================================================================================= 
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery()
    	End If
    End If
    
End Sub

'========================================================================================================= 
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery															'☜: Query db data

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
    Call parent.FncFind(parent.C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
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

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If
    
    With frm1
<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
	If lgIntFlgMode = parent.OPMD_UMODE Then  
	
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>

		strVal = strVal & "&txtCreditGrp=" & Trim(.HtxtCreditGrp.value)
		strVal = strVal & "&txtCurrency=" & Trim(.HtxtCurrency.value)    

        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag

    Else

        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									<%'☜: 비지니스 처리 ASP의 상태 %>

		strVal = strVal & "&txtCreditGrp=" & Trim(.txtCreditGrp.value)
		strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)    
    
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
        
    End if
		
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True


End Function

'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode    

    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolBar("11000000000111")							'⊙: 버튼 툴바 제어 
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus		
    End If     

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>여신현황</font></td>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>여신관리그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="여신관리그룹"  NAME="txtCreditGrp" SIZE="10" MAXLENGTH="5" tag="11XNXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCreditGrp() ">&nbsp;<INPUT TYPE=TEXT NAME="txtCreditGrpNm" SIZE="20" tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>화폐</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCurrency" ALT="화폐" TYPE="Text" MAXLENGTH="3" SIZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCurrency()"></TD>
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
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/s1411qa1_vaSpread1_vspdData.js'></script>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">여신관리그룹등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="HtxtCreditGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtCurrency" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
