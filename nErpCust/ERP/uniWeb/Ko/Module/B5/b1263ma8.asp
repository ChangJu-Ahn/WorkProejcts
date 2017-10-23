<%@ LANGUAGE="VBSCRIPT" %>
<%
'************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : B1263MA8
'*  4. Program Name         : 사업자이력조회 
'*  5. Program Desc         : 사업자이력조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Sonbumyeol
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : -2000/04/29 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO변환 
'*                            -2002/12/06 : UI성능향상(include) 반영 강준구 
'*                            -2002/12/11 : UI성능향상(include) 다시 반영 강준구 
'**************************************************************************************
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

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit				'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" --> 
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 

Dim lgMark                                                  <%'☜: 마크                                  %>
'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "b1263mb8.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID = "b1263ma1"

Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value

Dim IsOpenPop 
Dim lsValidDate             

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgSortKey   = 1
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtConBp_cd.focus  
End Sub

'========================================================================================================= 
<% '== 조회,출력 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("B1263MA8","S","A","V20021106", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub
'========================================================================================================= 
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'========================================================================================================= 
Sub SetSpreadColor(ByVal lRow)
End Sub

'========================================================================================================= 
Function OpenConBp_cd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처"    <%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"				<%' TABLE 명칭 %>

	arrParam(2) = Trim(frm1.txtConBp_cd.value)	<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "거래처"					<%' TextBox 명칭 %>
	
    arrField(0) = "BP_CD"						<%' Field명(0)%>
    arrField(1) = "BP_NM"						<%' Field명(1)%>
    
    arrHeader(0) = "거래처"					<%' Header명(0)%>
    arrHeader(1) = "거래처약칭"				<%' Header명(1)%>
    
	frm1.txtConBp_cd.focus

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
    
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConBpCode(arrRet)
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
Function SetConBpCode(Byval arrRet)

	frm1.txtConBp_cd.value = arrRet(0) 
	frm1.txtConBp_nm.value = arrRet(1)   

End Function

'========================================================================================================= 
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>

	Dim strTemp, arrVal

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If Kubun = 1 Then

		If lsValidDate = "" Then
			lsValidDate = "<%=GetSvrDate%>"
		End If
		
		WriteCookie CookieSplit , frm1.txtConBp_cd.value & parent.gRowSep & frm1.txtConBp_nm.value & parent.gRowSep & lsValidDate

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
		
		If strTemp = "" then Exit Function
		
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" then Exit Function
		
		frm1.txtConBp_cd.value =  arrVal(0)
		frm1.txtConBp_nm.value =  arrVal(1)
		
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

    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	
	'----------  Coding part  -------------------------------------------------------------
	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	Call CookiePage(0)
	
	frm1.txtConBp_cd.focus
	
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",1) ' 1
	lsValidDate=frm1.vspdData.Text
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Sub

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'========================================================================================================= 
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
    
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then  Exit Sub
    
	If CheckRunningBizProcess = True Then	   Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End If
    
End Sub

'========================================================================================================= 
Sub txtConValidFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConValidFromDt.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtConValidFromDt.Focus
	End If
End Sub
Sub txtConValidToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConValidToDt.Action = 7
		Call SetFocusToDocument("M")   
		Frm1.txtConValidToDt.Focus
	End If
End Sub

'========================================================================================================= 
Sub txtConValidFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtConValidToDt_KeyDown(KeyCode, Shift)
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

	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtConValidFromDt, frm1.txtConValidToDt) = False Then Exit Function

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
	Call SetDefaultVal
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data

    FncQuery = True																'⊙: Processing is OK

End Function

'========================================================================================================= 
Function FncPrint() 
    ggoSpread.Source = frm1.vspdData
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vb
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================= 
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	    
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If


	Dim strVal
    
    With frm1

		If lgIntFlgMode = parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
		    strVal = strVal & "&txtConBp_cd=" & Trim(.txtHBpCode.value)
		    strVal = strVal & "&txtConValidFromDt=" & Trim(.txtHValidFDt.value)
		    strVal = strVal & "&txtConValidToDt=" & Trim(.txtHValidTDt.value)		    
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtConBp_cd=" & Trim(.txtConBp_cd.value)
			strVal = strVal & "&txtConValidFromDt=" & Trim(.txtConValidFromDt.Text)
			strVal = strVal & "&txtConValidToDt=" & Trim(.txtConValidToDt.Text)
		End If
		<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>	
			strVal = strVal & "&lgPageNo="       & lgPageNo                '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================================= 
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
	lgIntFlgMode = parent.OPMD_UMODE                   'Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    Call SetToolBar("11000000000111")										'⊙: 버튼 툴바 제어 
    
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    End if      

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사업자이력조회</font></td>
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
                  <TD CLASS="TD5" NOWRAP>거래처</TD>
                  <TD CLASS="TD6"><INPUT NAME="txtConBp_cd" ALT="거래처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBp_cd">&nbsp;<INPUT NAME="txtConBp_nm" TYPE="Text" SIZE=25 tag="14"></TD>
                  <TD CLASS="TD5" NOWRAP>적용일</TD>
									<TD CLASS="TD6">
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
                          <script language =javascript src='./js/b1263ma8_I901444697_txtConValidFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
                          <script language =javascript src='./js/b1263ma8_I593622535_txtConValidToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
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
										<script language =javascript src='./js/b1263ma8_vaSpread1_vspdData.js'></script>
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
          <TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">사업자이력등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		            FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<INPUT TYPE=HIDDEN NAME="txtHBpCode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHValidFDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHValidTDt" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>

</BODY>
</HTML> 

