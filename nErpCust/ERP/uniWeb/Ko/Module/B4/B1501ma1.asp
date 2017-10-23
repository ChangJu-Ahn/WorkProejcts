
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Calendar수정)
'*  3. Program ID           : B1501ma1.asp
'*  4. Program Name         : B1501ma1.asp
'*  5. Program Desc         : 칼렌다수정 
'*  6. Modified date(First) : 2000/09/15
'*  7. Modified date(Last)  : 2002/01/23
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<STYLE TYPE="text/css">
	.Header {height:24; font-weight:bold; text-align:center; color:darkblue}
	.Day {height:22;cursor:Hand;
		font-size:17; font-weight:bold; Border:0; text-align:right}
	.DummyDay {height:22;cursor:;
		font-size:12; font-weight:; Border:0; text-align:right}
</STYLE>
<MAP NAME="CalButton">
	<AREA SHAPE=RECT COORDS="1, 1, 20, 20" ALT="Year -" onClick="ChangeMonth(-12)">
	<AREA SHAPE=RECT COORDS="20, 1, 40, 20" ALT="Month -" onClick="ChangeMonth(-1)">
	<AREA SHAPE=RECT COORDS="40, 1, 60, 20" ALT="Month +" onClick="ChangeMonth(1)">
	<AREA SHAPE=RECT COORDS="60, 1, 80, 20" ALT="Year +" onClick="ChangeMonth(12)">
</MAP>

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "B1501mb1.asp"											'☆: 비지니스 로직 ASP명 

Const CChnageColor = "#f0fff0"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

Dim lgLastDay
Dim lgStartIndex
Dim lgArrDate(31, 3)

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------

	Dim iRow, iCol
	For iRow = 1 To 6
		For iCol = 1 To 7
			If frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = CChnageColor Then
				frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = "white"
				frm1.txtDate((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
				frm1.txtDesc((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
			End If
		Next
	Next
End Sub

Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "B","NOCOOKIE","MA") %>
End Sub


Sub SetDefaultVal()

    Dim strYear
    Dim strMonth
    Dim strDay
    
    Call ExtractDateFrom("<%= GetSvrDate %>",parent.gServerDateFormat , parent.gServerDateType      ,strYear,strMonth,strDay)

	frm1.txtYymm.Year  = strYear
	frm1.txtYymm.Month = strMonth

    Frm1.txtYymm.focus 
    
End Sub

Sub Form_Load()

    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                            <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)                            
    
    Call ggoOper.FormatDate(frm1.txtYymm, parent.gDateFormat, 2)
    Call InitVariables																'⊙: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("1100100000000011")
    Call SetDefaultVal
    
    Call FncQuery()
End Sub

Sub DescChange(iDate)
	Dim strDesc
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If
	
	Call SetChange(iDate)

	strDesc = frm1.txtDesc(index).value
	frm1.txtDesc(index).value = ""
	
	frm1.txtDesc(index).value = strDesc
	frm1.txtDesc(index).title = strDesc
End Sub

Sub HoliChange(iDate)
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If

	Call SetChange(iDate)
	
	If frm1.txtHoli(index).value = "H" Then
		If (index+1) Mod 7 = 0 Then
			frm1.txtDate(index).style.color = "blue"
		Else
			frm1.txtDate(index).style.color = "black"
		End If
		frm1.txtHoli(index).value = "D"
	Else
		frm1.txtDate(index).style.color = "red"
		frm1.txtHoli(index).value = "H"
	End if
End Sub

Sub SetChange(iDate)
	Dim index
	index = iDate - 1

	lgBlnFlgChgValue = True
	
	frm1.All.tblCal.Rows(Int((index+7)/7)).Cells(index Mod 7).Style.backgroundColor = CChnageColor
	frm1.txtDate(index).Style.backgroundColor = CChnageColor
	frm1.txtDesc(index).Style.backgroundColor = CChnageColor
End Sub

Sub ChangeMonth(i)
    Dim strVal
    Dim dtDate
    Dim IntRetCD

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Sub
		End If
    End If

    Call InitVariables	'⊙: Initializes local global variables
	
	On Error Resume Next
	Err.Clear
	
    dtDate = CDate(frm1.hYear.value & "-" & frm1.hMonth.value & "-" & "01")

    If Err.Number <> 0 Then                         'Check if there is retrived data        
        Err.Clear
		Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Sub
    End If

	dtDate = DateAdd("m", i, dtDate)
	
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
    strVal = strVal & "&txtYear=" & Year(dtDate)							'☆: 조회 조건 데이타 
    strVal = strVal & "&txtMonth=" & Month(dtDate)							'☆: 조회 조건 데이타 

	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)
End Sub

Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                        '⊙: Processing is NG
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function							<%'Query db data%>
       
    FncQuery = True																'⊙: Processing is OK
        
End Function

Function FncSave() 
        
    FncSave = False                                                         '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                            '⊙: No data changed!!
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
       
    End If
    
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

Function FncPrint() 
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			<%'⊙: "Will you destory previous data"%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery()
	Dim strYear, strMonth
    
    strYear = Frm1.txtYymm.Year
    strMonth = Frm1.txtYymm.Month
    
    DbQuery = False                                                         '⊙: Processing is NG

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
    strVal = strVal & "&txtYear=" & strYear	'☆: 조회 조건 데이타 
    strVal = strVal & "&txtMonth=" & strMonth	'☆: 조회 조건 데이타 

	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbQuery = True                                                          '⊙: Processing is NG

End Function

Function DbQueryOk()													'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
End Function

Function DbSave() 

    Err.Clear																	'☜: Protect system from crashing

	DbSave = False																'⊙: Processing is NG

	frm1.txtMode.value = parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode
	frm1.txtUpdtUserId.value = parent.gUsrID
	
	If Not chkField(Document, "2") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
	Call LayerShowHide(1)
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

    Call InitVariables()

End Function

Sub txtYymm_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYymm.Action = 7
        Call SetFocusToDocument("M") 
        frm1.txtYymm.Focus
    End If
End Sub

Sub txtYymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>칼렌다수정</font></td>
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
					<TD>
						<TABLE ID="tbTitle" WIDTH=100% BORDER=0 CELLSPACING=0 ALIGN="center">
							<TR>
								<TD WIDTH=* Style="font-size:20; font-weight:bold; text-align:center" NoWrap>&nbsp;</TD>
								<TD Width=10><IMG SRC="../../../CShared/image/CalButton.gif" Width=80 HEIGHT=20 style="cursor:Hand" ISMAP USEMAP="#CalButton"></IMG></TD><TD Width=5>&nbsp;</TD>
								<TD WIDTH=10><script language =javascript src='./js/b1501ma1_fpDateTime_txtYymm.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD>
						<TABLE ID="tblCal" WIDTH=100% HEIGHT=100% BORDER=1 CELLSPACING=0 CELLPADDING=0 ALIGN="center">
							<THEAD CLASS="Header">
								<TR>
									<TD>일요일</TD>
									<TD>월요일</TD>
									<TD>화요일</TD>
									<TD>수요일</TD>
									<TD>목요일</TD>
									<TD>금요일</TD>
									<TD>토요일</TD>
					            </TR>
				        	</THEAD>
							<TBODY>
<%
Dim i, j, k
k = 1
For i=1 To 6
%>
					            <TR>
<%
	For j=1 To 7
%>
									<TD ALIGN="Center">
										<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 ALIGN="Center">
											<TR>
												<TD ALIGN="Left">
													<INPUT type="hidden" name="txtHoli" size=1 maxlength=1 disabled>
													<INPUT type="text" name="txtDate" class="DummyDay" size=2 maxlength=2  
														tabindex=-1 readonly disabled onclick="HoliChange(<%=k%>)">
												</TD>
											</TR>
											<TR>
												<TD ALIGN="Left">
													<INPUT type="text" name="txtDesc" MaxLength=30 Style="Width:100%;Border:0;text-align:center" disabled tag=2 onchange="DescChange(<%=k%>)" >
												</TD>
											</TR>											
										</TABLE>
									</TD>
<%
		k = k + 1
	Next
%>
								</TR>
<%
Next
%>
							</TBODY>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="B1501mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hYear" tag="24">
<INPUT TYPE=HIDDEN NAME="hMonth" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

