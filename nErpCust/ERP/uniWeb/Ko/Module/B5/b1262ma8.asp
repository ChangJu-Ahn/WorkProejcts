<%@ LANGUAGE="VBSCRIPT" %>
<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 기준정보 
'*  3. Program ID           : B1262MA8
'*  4. Program Name         : 거래처형태 조회 
'*  5. Program Desc         : 거래처형태 조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : kim hyung suk
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/29 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO변환 
'*                            -2002/12/05 : UI성능향상(include) 반영 강준구 
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
Dim IsOpenPop
Dim arrParam	

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "b1262mb8.asp"												'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_JUMP_ID1 = "b1262ma1"
Const BIZ_PGM_JUMP_ID2 = "b1262ma2"
Const BIZ_PGM_JUMP_ID3 = "b1261ma1"

Const C_MaxKey          = 4                                    '☆☆☆☆: Max key value

<% '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- %>
Dim lsPartFtnCd
Dim lsPartBpCd
Dim lsPartBpNm

Dim lsSplict

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
	lgPageNo         = ""
    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtBp_cd.focus  
End Sub

'========================================================================================================= 
<% '== 조회,출력 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("B1262MA8","S","A","V20021106", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub
'========================================================================================================= 
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================= 
Function OpenBp_cd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처"          <%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"				<%' TABLE 명칭 %>

	arrParam(2) = Trim(frm1.txtBp_cd.value)		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = ""							<%' Where Condition%>
	arrParam(5) = "거래처"					<%' TextBox 명칭 %>
	
    arrField(0) = "BP_CD"						<%' Field명(0)%>
    arrField(1) = "BP_NM"						<%' Field명(1)%>
    
    arrHeader(0) = "거래처"					<%' Header명(0)%>
    arrHeader(1) = "거래처약칭"				<%' Header명(1)%>
    
    frm1.txtBp_cd.focus 
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
    
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBpCode(arrRet)
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
Function SetBpCode(Byval arrRet)

	frm1.txtBp_cd.value = arrRet(0) 
	frm1.txtBp_nm.value = arrRet(1)   

End Function

'========================================================================================================= 
Function LoadPageCheck_S()

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If lsSplict = "CS" or lsSplict = "S" Then
		Call CookiePage(1)
		PgmJump(BIZ_PGM_JUMP_ID1)
	Else
		Call DisplayMsgBox("126234","x","x","x")
        'MsgBox "매입처형태등록 화면으로 이동할 수 없습니다. 매입/매출구분을 확인하십시오.", vbExclamation, "uniERP(Warning)"
	End If   

End Function

'========================================================================================================= 
Function LoadPageCheck_C()

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If lsSplict = "CS" or lsSplict = "C" Then
		Call CookiePage(1)
		PgmJump(BIZ_PGM_JUMP_ID2)
	Else
		Call DisplayMsgBox("126235","x","x","x")
        'MsgBox "매출처형태등록 화면으로 이동할 수 없습니다. 매입/매출구분을 확인하십시오.", vbExclamation, "uniERP(Warning)"
	End IF   

End Function

'========================================================================================================= 
Function CookiePage(ByVal Kubun)

	on error resume next
	
	Const CookieSplit = 4877
	
	Dim strTemp, arrVal

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If Kubun = 1 Then

		WriteCookie CookieSplit , frm1.txtBp_cd.value & parent.gRowSep & frm1.txtBp_nm.value _
			& parent.gRowSep & lsPartBpCd & parent.gRowSep & lsPartBpNm & parent.gRowSep & lsPartFtnCd

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
		
		If strTemp = "" then Exit Function
		
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" then Exit Function
		
		frm1.txtBp_cd.value =  arrVal(0)
		frm1.txtBp_nm.value =  arrVal(1)
		
		If Err.number <> 0 then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit function
		End If
		
		Call MainQuery()
		
		WriteCookie CookieSplit , ""

	End IF

End Function

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														

    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
	
	'----------  Coding part  -------------------------------------------------------------
	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	
	Call CookiePage(0)
	
    frm1.txtBp_cd.focus
    	
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
	lsSplict=frm1.vspdData.Text

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",2) ' 6
	lsPartBpNm=frm1.vspdData.Text

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",3) ' 3
	lsPartFtnCd=frm1.vspdData.Text

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",4) ' 5
	lsPartBpCd=frm1.vspdData.Text    
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
    
    If OldLeft <> NewLeft Then Exit Sub
   
	If CheckRunningBizProcess = True Then   Exit Sub

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
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

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
	Call parent.FncFind(parent.C_MULTI, False)
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
    
        
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtBp_cd=" & Trim(.HBp_cd.value)
			
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtBp_cd=" & Trim(.txtBp_cd.value)
			
		End if
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
                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>거래처형태조회</font></td>
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
									<TD CLASS="TD6" NOWRAP>
                    <input NAME="txtBp_cd" TYPE="Text" MAXLENGTH="10" tag="12XXXU" size="10" Alt="거래처"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgBp_cd" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenBp_cd()"> 
										<input NAME="txtBp_nm" TYPE="Text" tag="14" size="30"></TD>
									<TD CLASS="TDT"></TD>
									<TD CLASS="TD6"></TD>
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
										<script language =javascript src='./js/b1262ma8_I686355620_vspdData.js'></script>
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
          <TD WIDTH=* ALIGN=RIGHT><a ONCLICK="VBSCRIPT:LoadPageCheck_C()"> 매출거래처형태등록</a>&nbsp;|&nbsp;<a ONCLICK="VBSCRIPT:LoadPageCheck_S()">매입거래처형태등록</a>&nbsp;|&nbsp;<a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID3)" ONCLICK="VBSCRIPT:CookiePage 1">거래처등록</a></TD>
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

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="HBp_cd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>
</BODY>
</HTML> 

