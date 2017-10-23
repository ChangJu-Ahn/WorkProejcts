<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1413QA1
'*  4. Program Name         : 담보현황조회 
'*  5. Program Desc         : 담보현황조회 
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

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
Dim prDBSYSDate
Dim EndDate ,StartDate
prDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAToB(prDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
StartDate = UniDateAdd("m", -1, EndDate,parent.gDateFormat)
'--------------- 개발자 coding part(실행로직,End)-------------------------------------------------------------

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s1413qb1.asp"
Const BIZ_PGM_JUMP_ID   = "s1413ma1"				  	       '☆: 비지니스 로직 ASP명 

Const gstrColletralTypeMajor = "S0002"
Const gstrDelTypeMajor = "S0003"
Const C_MaxKey          = 15                                   '☆☆☆☆: Max key value
Dim lsColletralNo                                             '☆: Jump시 Cookie로 보낼 Grid value

'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode

End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBpCd.focus
	frm1.txtAsignFrDt.text = StartDate
	frm1.txtAsignToDt.text = EndDate

End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub


'========================================================================================================= 
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S1413QA1","S","A","V20030318", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
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
	Function OpenBizPartner()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If lgIsOpenPop = True Then Exit Function

		lgIsOpenPop = True

		arrParam(0) = "고객"							<%' 팝업 명칭 %>
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtBpCd.value)				<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				<%' Where Condition%>
		arrParam(5) = "고객"							<%' TextBox 명칭 %>

		arrField(0) = "BP_CD"								<%' Field명(0)%>
		arrField(1) = "BP_NM"								<%' Field명(1)%>

		arrHeader(0) = "고객"							<%' Header명(0)%>
		arrHeader(1) = "고객명"						<%' Header명(1)%>

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False
		
		frm1.txtBpCd.focus 

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetBizPartner(arrRet)
		End If
	End Function

'========================================================================================================= 
	Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If lgIsOpenPop = True Then Exit Function

		lgIsOpenPop = True

		arrParam(0) = strPopPos								<%' 팝업 명칭 %>
		arrParam(1) = "B_Minor"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(strMinorCD)						<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		<%' Where Condition%>
		arrParam(5) = strPopPos								<%' TextBox 명칭 %>

		arrField(0) = "Minor_CD"							<%' Field명(0)%>
		arrField(1) = "Minor_NM"							<%' Field명(1)%>

		arrHeader(0) = strPopPos							<%' Header명(0)%>
		arrHeader(1) = strPopPos & "명"							<%' Header명(1)%>

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False
		
		Select Case strMajorCd
			Case gstrColletralTypeMajor
				frm1.txtColletralType.focus 
			Case gstrDelTypeMajor
				frm1.txtDelType.focus 
			Case Else
		End Select		

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetMinorCd(strMajorCd, arrRet)
		End If
	End Function

'========================================================================================================= 
	Function OpenSalesGroup()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If lgIsOpenPop = True Then Exit Function

		lgIsOpenPop = True

		arrParam(0) = "영업그룹"						<%' 팝업 명칭 %>
		arrParam(1) = "B_SALES_GRP"							<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtSalesGroup.value)		<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
		arrParam(5) = "영업그룹"						<%' TextBox 명칭 %>

		arrField(0) = "SALES_GRP"							<%' Field명(0)%>
		arrField(1) = "SALES_GRP_NM"						<%' Field명(1)%>

		arrHeader(0) = "영업그룹"						<%' Header명(0)%>
		arrHeader(1) = "영업그룹명"						<%' Header명(1)%>

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False
		
	    frm1.txtSalesGroup.focus 

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetSalesGroup(arrRet)
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
	Function SetBizPartner(arrRet)
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
	End Function

'========================================================================================================= 
	Function SetMinorCd(strMajorCd, arrRet)
		Select Case strMajorCd
			Case gstrColletralTypeMajor
				frm1.txtColletralType.value = arrRet(0)
				frm1.txtColletralTypeNm.value = arrRet(1)
			Case gstrDelTypeMajor
				frm1.txtDelType.value = arrRet(0)
				frm1.txtDelTypeNm.value = arrRet(1)
			Case Else
		End Select
	End Function

'========================================================================================================= 
	Function SetSalesGroup(arrRet)
		frm1.txtSalesGroup.value = arrRet(0)
		frm1.txtSalesGroupNm.value = arrRet(1)
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

		WriteCookie CookieSplit , lsColletralNo		<% 'Jump로 화면을 이동할때 필요한 Cookie 변수정의 %>
		
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
		frm1.txtSalesGroup.value = arrVal(0)
		frm1.txtSalesGroupNm.value = arrVal(1)
		frm1.txtColletralType.value = arrVal(2)
		frm1.txtColletralTypeNm.value = arrVal(3)
		frm1.txtAsignFrDt.text = arrVal(4)
		frm1.txtAsignToDt.text = arrVal(4)
		frm1.txtDelType.value = arrVal(5)
		frm1.txtDelTypeNm.value = arrVal(6)
		frm1.txtBpCd.value = arrVal(7)
		frm1.txtBpNm.value = arrVal(8)  
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
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
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
	frm1.vspdData.col = GetKeyPos("A",3) 
	lsColletralNo=frm1.vspdData.Text
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
    
    If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		If lgStrPrevKey <> "" Then		
			If CheckRunningBizProcess Then	Exit Sub				
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery()
		End if
	End if	    


End Sub
'========================================================================================================= 
	Sub btnColletralTypeOnClick()
		Call OpenMinorCd(frm1.txtColletralType.value, frm1.txtColletralTypeNm.value, "담보유형", gstrColletralTypeMajor)
	End Sub

'========================================================================================================= 
	Sub btnDelTypeOnClick()
		Call OpenMinorCd(frm1.txtDelType.value, frm1.txtDelTypeNm.value, "해지구분", gstrDelTypeMajor)
	End Sub

'========================================================================================================= 
	Sub rdoColStateFlg1_OnClick()
		frm1.txtRadio.value = frm1.rdoColStateFlg1.value
	End Sub

	Sub rdoColStateFlg2_OnClick()
		frm1.txtRadio.value = frm1.rdoColStateFlg2.value
	End Sub

	Sub rdoColStateFlg3_OnClick()
		frm1.txtRadio.value = frm1.rdoColStateFlg3.value
	End Sub

'========================================================================================================= 
Sub txtAsignFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtAsignFrDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtAsignFrDt.Focus
	End If
End Sub
Sub txtAsignToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtAsignToDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtAsignToDt.Focus
	End If
End Sub

'========================================================================================================= 
Sub txtAsignFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtAsignToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'========================================================================================================= 
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
 
    '** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 커야 할때 **
	If ValidDateCheck(frm1.txtAsignFrDt, frm1.txtAsignToDt) = False Then Exit Function

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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
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
	
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If
    
    With frm1
		
		If lgIntFlgMode = parent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtBpCd=" & Trim(.HBpCd.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.HSalesGroup.value)
			strVal = strVal & "&txtColletralType=" & Trim(.HColletralType.value)
			strVal = strVal & "&txtAsignFrDt=" & Trim(.HAsignFrDt.value)
			strVal = strVal & "&txtAsignToDt=" & Trim(.HAsignToDt.value)
			strVal = strVal & "&txtRadio=" & Trim(.HRadio.value)
			strVal = strVal & "&txtDelType=" & Trim(.HDelType.value)
		Else
			strVal = BIZ_PGM_ID & "?txtBpCd=" & Trim(.txtBpCd.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
			strVal = strVal & "&txtColletralType=" & Trim(.txtColletralType.value)
			strVal = strVal & "&txtAsignFrDt=" & Trim(.txtAsignFrDt.text)
			strVal = strVal & "&txtAsignToDt=" & Trim(.txtAsignToDt.text)
			strVal = strVal & "&txtRadio=" & Trim(.txtRadio.value)
			strVal = strVal & "&txtDelType=" & Trim(.txtDelType.value)
		End If
		
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
	
	lgIntFlgMode = parent.OPMD_UMODE
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>담보현황</font></td>
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
									<TD CLASS=TD5 NOWRAP>고객</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10  MAXLENGTH=10 TAG="11XXXU" ALT="고객"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="VBScript:OpenBizPartner()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 TAG="14">
									</TD>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="VBScript:OpenSalesGroup()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>담보유형</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtColletralType" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="담보유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnColletralType" align=top TYPE="BUTTON" ONCLICK="VBScript:btnColletralTypeOnClick()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtColletralTypeNm" SIZE=20 TAG="14">
									</TD>
									<TD CLASS=TD5 NOWRAP>설정일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s1413qa1_fpDateTime1_txtAsignFrDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s1413qa1_fpDateTime2_txtAsignToDt.js'></script>
									</TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>담보상태</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoColStateFlg" TAG="11X" VALUE="A" CHECKED ID="rdoColStateFlg1" ><LABEL FOR="rdoColStateFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoColStateFlg" TAG="11X" VALUE="N" ID="rdoColStateFlg2" ><LABEL FOR="rdoColStateFlg2">설정</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoColStateFlg" TAG="11X" VALUE="Y" ID="rdoColStateFlg3" ><LABEL FOR="rdoColStateFlg3">해지</LABEL>			
									</TD>
									<TD CLASS=TD5 NOWRAP>해지구분</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDelType" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="해지구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDelType" align=top TYPE="BUTTON" ONCLICK="VBScript:btnDelTypeOnClick()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtDelTypeNm" SIZE=20 TAG="14">
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
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/s1413qa1_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">담보등록</a></TD>
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

<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="14">
<INPUT TYPE=HIDDEN NAME="HBpCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="HSalesGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="HColletralType" TAG="24">
<INPUT TYPE=HIDDEN NAME="HAsignFrDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="HAsignToDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="HRadio" TAG="24">
<INPUT TYPE=HIDDEN NAME="HDelType" TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>

</BODY>
</HTML>

