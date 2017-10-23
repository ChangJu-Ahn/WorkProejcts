<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Basic Info.
'*  3. Program ID           : A2108MA1
'*  4. Program Name         : 전표마감등록 
'*  5. Program Desc         : Register of G/L Closing
'*  6. Component List       : A2108M
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/08/09
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Shin Myoung_Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->
<STYLE TYPE="text/css">
	.Header {height:24; font-weight:bold; text-align:center; color:darkblue}
	.Day {width:25;height:25;font-size:17; font-weight:bold; Border:0; text-align:right}
	.DummyDay {width:25;height:22;font-size:12; font-weight:; Border:0; text-align:right}
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

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit								'☜: indicates that All variables must be declared in advance

'==========================================================================================================
<%
Dim SvrDate
SvrDate = GetSvrDate
%>

Const BIZ_PGM_ID = "a2108mb1.asp"											'☆: 비지니스 로직 ASP명 

Const CChnageColor = "#f0fff0"

'=========================================================================================================

Dim lgBlnFlgChgValue
Dim lgIntGrpCount
Dim lgIntFlgMode

Dim lgNextNo
Dim lgPrevNo						
 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim lgLastDay
Dim lgStartIndex
Dim lgArrDate(31, 3)

'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
	Dim iRow, iCol
	For iRow = 1 To 6
		For iCol = 1 To 7
			If frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = CChnageColor Then
				frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = "white"
				frm1.txtDate((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
				frm1.txtT((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
				frm1.txtG((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
				frm1.txtDesc((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
			End If
		Next
	Next

End Sub


'========================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

'========================================================================================
Sub InitComboBox()
	Dim i, ii

	Dim strYear, strMonth, strDay,strDate
    strDate               = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	For i = strYear - 10 To strYear + 10
		Call SetCombo(frm1.cboYear, i, i)
	Next

	For i=1 To 12
		ii = Right("0" & i, 2)
		Call SetCombo(frm1.cboMonth, ii, ii)
	Next

    frm1.cboYear.value = strYear
    frm1.cboMonth.value = strMonth
End Sub

'========================================================================================
Sub SetDefaultVal()
'	On Error Resume Next
End Sub

'========================================================================================
Sub Form_Load()
    Call InitVariables
    Call LoadInfTB19029
    Call InitComboBox
    Call SetToolbar("1100100000001111")
    Call SetDefaultVal
	Call FncQuery()
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
'	NAME :  FlagTChange(iDate)
'	기능 :  Temp G/L Change
'==========================================================================================
Sub FlagTChange(iDate)
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If

	If frm1.txtFlgT(index).value = "C" Then
		frm1.txtT(index).style.color = "silver"
		frm1.txtFlgT(index).value = "O"
	Else
		frm1.txtT(index).style.color = "blue"
		frm1.txtFlgT(index).value = "C"
	End if

	Call SetChange(iDate)
End Sub

'==========================================================================================
'	NAME :  FlagGChange(iDate)
'	기능 :  G/L Change
'==========================================================================================
Sub FlagGChange(iDate)
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If

	If frm1.txtFlgG(index).value = "C" Then
		frm1.txtG(index).style.color = "silver"
		frm1.txtFlgG(index).value = "O"
	Else
		frm1.txtG(index).style.color = "red"
		frm1.txtFlgG(index).value = "C"
	End if

	Call SetChange(iDate)
End Sub

'==========================================================================================
' NAME : SetChange(iDate)
' 기능 : Change Setting
'==========================================================================================
Sub SetChange(iDate)
	Dim index
	index = iDate - 1

	lgBlnFlgChgValue = True

	frm1.All.tblCal.Rows(Int((index+7)/7)).Cells(index Mod 7).Style.backgroundColor = CChnageColor
	frm1.txtDate(index).Style.backgroundColor = CChnageColor
	frm1.txtT(index).Style.backgroundColor = CChnageColor
	frm1.txtG(index).Style.backgroundColor = CChnageColor
	frm1.txtDesc(index).Style.backgroundColor = CChnageColor
End Sub

'==========================================================================================
' NAME : ChangeMonth(i)
' 기능 : Change Month
'==========================================================================================
Sub ChangeMonth(i)
    Dim strVal
    Dim dtDate
    Dim IntRetCD
    Dim strYear
	Dim strMonth
    Dim strDay
    
	'##########################################
	'TEST
	'IF lgBlnFlgChgValue = TRUE THEN
	'	MSGBOX "TRUE"
	'ELSE
	'	MSGBOX "FALSE"
	'END IF
	'##########################################
	If lgBlnFlgChgValue = True Then
		'MSGCODE(900013) => 데이타가 변경되었습니다. 조회하시겠습니까?
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X") '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Sub
		End If
	End If

	Call InitVariables				'☜: Initializes local global variables

	'On Error Resume Next
	Err.Clear

	dtDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, frm1.txtYear.value, frm1.txtMonth.value, "01")
    
    If Err.Number <> 0 Then         '☜: Check if there is retrived data
		Call DisplayMsgBox("900002","X","X","X")
        Exit Sub
    End If
	
	dtDate = UNIDateAdd("M", i, dtDate, Parent.gDateFormat)

	Call ExtractDateFrom(dtDate,Parent.gDateFormat,Parent.gComDateType,strYear,strMonth,strDay)
	frm1.cboYear.value = strYear
	frm1.cboMonth.value = strMonth

	Call DbQuery
End Sub

'========================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")		'☜: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Call InitVariables
    Call DbQuery

    FncQuery = True
End Function

'========================================================================================
Function FncSave()
    Dim IntRetCD

    FncSave = False

    Err.Clear
    'On Error Resume Next

    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                      '⊙: No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    'If Not chkField(Document, "2") Then									'⊙: Check contents area
    '   Exit Function
    'End If

    '-----------------------
    'Save function call area
    '-----------------------
    CAll DbSave
    FncSave = True
End Function

'========================================================================================
Function FncCopy()

End Function

'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================
Function FncCancel()
    On Error Resume Next
End Function

'========================================================================================
Function FncInsertRow()
     On Error Resume Next
End Function

'========================================================================================
Function FncDeleteRow()
    On Error Resume Next
End Function

'========================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================
Function FncPrev()
	On Error Resume Next
End Function

'========================================================================================
Function FncNext()
	On Error Resume Next
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

		' 변경된 내용이 있는지 확인한다.
		If lgBlnFlgChgValue = True Then
		     IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
		     If IntRetCD = vbNo Then
		         Exit Function
		    End If
		End If

    FncExit = True
End Function

'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜:
    strVal = strVal & "&txtYear=" & Trim(frm1.cboYear.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtMonth=" & Trim(frm1.cboMonth.Value)				'☆: 조회 조건 데이타 
	strVal = strVal & "&htxtTempGl=" & Trim(frm1.htxtTempGl.value)
	strVal = strVal & "&htxtGl=" & Trim(frm1.htxtGl.value)

	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
End Function

'========================================================================================
Function DbSave()

    Err.Clear

	DbSave = False

	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtFlgMode.value = lgIntFlgMode

	Call LayerShowHide(1)
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True

End Function

'========================================================================================
Function DbSaveOk()	
    Call InitVariables
	'필요성 
	Call FncQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<!-- 타이틀부분전표마감등록-->
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>전표마감일등록</font></td>
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
			<TABLE <%=LR_SPACE_TYPE_20%> >
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE ID="tbTitle" <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD WIDTH=* Style="font-size:20; font-weight:bold; text-align:center" NoWrap>&nbsp;</TD>
									<TD Width=10><IMG SRC="../../../CShared/image/CalButton.gif" Width=80 HEIGHT=20 style="cursor:Hand" ISMAP USEMAP="#CalButton"></IMG></TD>
									<TD Width=5>&nbsp;</TD>
									<TD WIDTH=10>
										<SELECT Name="cboYear" STYLE="WIDTH=60"></SELECT>
									</TD>
									<TD WIDTH=10>
										<SELECT Name="cboMonth" STYLE="WIDTH=40"></SELECT>
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
					<TD>
						<TABLE ID="tblCal" BORDER=1 cellpadding="0" <%=LR_SPACE_TYPE_20%>>
							<THEAD CLASS="Header">
								<TR>
									<TD width="10%">일요일</TD>
									<TD width="10%">월요일</TD>
									<TD width="10%">화요일</TD>
									<TD width="10%">수요일</TD>
									<TD width="10%">목요일</TD>
									<TD width="10%">금요일</TD>
									<TD width="10%">토요일</TD>
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
									<TD ALIGN="Center" >
										<TABLE WIDTH=100% cellpadding="0" CELLSPACING=0 ALIGN="Center">
											<TR>
												<TD Align="Left">
													<INPUT type="text" name="txtDate" class="DummyDay" size=2 maxlength=2 tabindex=-1 readonly>
												</TD>
												<TD >
													<INPUT type="hidden" name="txtFlgT" size=4 maxlength=4 disabled  TABINDEX = "-1">
													<INPUT type="text" name="txtT" style="width:40;Border:0;text-align:center" size=4 readonly onclick="FlagTChange(<%=k%>)">
												</TD>
												<TD >
													<INPUT type="hidden" name="txtFlgG" size=4 maxlength=4 disabled TABINDEX = "-1">
													<INPUT type="text" name="txtG" style="width:40;Border:0;text-align:center" size=4 readonly onclick="FlagGChange(<%=k%>)">
												</TD>
											</TR>
										</TABLE>
										<INPUT type="text" name="txtDesc" Style="Width:100%;Border:0;text-align:center" tabindex=-1 readonly></INPUT>
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
		<!--표준 HEIGHT=20-->
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" src="../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtYear" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMonth" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtTempGl" tag="34" Value="결의" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtGl" tag="34" Value="회계" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>