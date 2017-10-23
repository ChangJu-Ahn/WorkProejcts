<!-- #Include file="../inc/CommResponse.inc" -->
<%
	SYear = Year(Now)
	SMonth = Month(Now)
	SDay = Day(Now)
%>
<HTML>
<HEAD>
<Link Rel="stylesheet" Type="Text/css" HREF="../inc/SheetStyle.css">
<STYLE TYPE="text/css">
	.CalTitle {font-size:20; font-weight:bold; text-align:center}
	.Header {height:20; font-weight:bold; text-align: center;background-color:linen; color:brown; valign:center}
	.Day {text-align:center; valign:center; background-color:LIGHTCYAN; color:darkblue}
	.SelDay {text-align:center; valign:center; background-color:silver; color:yellow}
	.DummyDay {text-align:center; valign:center; background-color:LIGHTCYAN; color:silver}
</STYLE>

<TITLE>Calendar</TITLE>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/event.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
<!--
Option Explicit

Dim Selector
Dim lgCurrCell
Dim lgCurrI, lgCurrJ
Dim lgYear, lgMonth, lgDay
Dim lgInputLock

Dim dtTemp

Err.Clear
On Error Resume Next
dtTemp = CDate(window.dialogArguments)

If Err.number = 0 Then
	lgYear = Year(dtTemp)
	lgMonth = Month(dtTemp)
	lgDay = Day(dtTemp)
Else
	lgYear = <%=SYear%>
	lgMonth = <%=SMonth%>
	lgDay = <%=SDay%>
End If
On Error Goto 0

lgInputLock = False

Sub CreateCal()
	Dim i, j 
	Dim CalRow, CalCol
	Dim CalCell
	Dim dtDate

	dtDate = CDate(lgYear & "-" & lgMonth & "-" & "1")
	
	<%' -- 1일 이전 데이타 클리어 --- %>
	For i = (WeekDay(dtDate) - 1) -1 To 0 Step -1
		Set CalCell = document.all.calendar.tBodies.dayList.rows(0).cells(i)

		dtDate = DateAdd("d", -1, dtDate)
		CalCell.innerText = CStr(Day(dtDate))
		CalCell.className = "DummyDay"
	Next
	
	dtDate = CDate(lgYear & "-" & lgMonth & "-" & "1")

	CalRow = 0
	CalCol = WeekDay(dtDate) - 1
	
	Do While CInt(Month(dtDate)) = CInt(lgMonth)
		Set CalCell = document.all.calendar.tBodies.dayList.rows(CalRow).cells(CalCol)
		
		CalCell.innerText = CStr(Day(dtDate))
		
		If CInt(CalCell.innerText) = CInt(lgDay) Then
			CalCell.className = "SelDay"
			Set lgCurrCell = CalCell
			
			lgCurrI = CalRow
			lgCurrJ = CalCol
		Else
			CalCell.className = "DAy"
		End If
		
		If WeekDay(dtDate) = 7 Then
			CalRow = CalRow + 1 : CalCol = 0
		Else
			CalCol = CalCol + 1
		End If
		
		dtDate = DateAdd("d", 1, dtDate)
	Loop
	
	<%' -- 마지막 날 이후 데이타 클리어 --- %>
	For j = CalRow to 5
		For i = 0 to 6
			If j = CalRow And i < CalCol Then
			Else
				Set CalCell = document.all.calendar.tBodies.dayList.rows(j).cells(i)

				CalCell.innerText = CStr(Day(dtDate))
				CalCell.className = "DummyDay"
				
				dtDate = DateAdd("d", 1, dtDate)
			End If
		Next
	Next

	Set CalCell = Nothing	
	lgInputLock = False
End Sub

Sub InitComboBox()
	Dim i, ii
	Dim oOption
	
	For i=lgYear-100 To lgYear+100
		Call SetCombo(cboYear, i, i)
	Next

	cboYear.remove 0
    cboYear.value = CStr(lgYear)
    
	For i=1 To 12
		ii = Right("0" & i, 2)
		Call SetCombo(cboMonth, ii, ii)
	Next

	cboMonth.remove 0
	cboMonth.value = Right("0" & lgMonth, 2)
End Sub

Sub window_onload
	Call GetGlobalVar()
	Selector = "N"

	Call InitComboBox
	Call SetDate(lgYear, lgMonth, lgDay)
	Call MM_preloadImages("../image/Query.gif","../image/OK.gif","../image/Cancel.gif")
End Sub

Sub window_onunload
	if Selector = "Y" then
		window.returnValue = txtDate.value
	else
		window.returnValue = ""
	end if
End Sub

Sub DayDblClick(i, j)
	Dim CalCell

	Set CalCell = document.all.calendar.tBodies.dayList.rows(i).cells(j)
	
	If Trim(CalCell.innertext) <> "" then 
		Call DayClick(i, j)
		
		Call btnOK_onclick
	End If

	Set CalCell = Nothing
End Sub

Sub DayClick(i, j)
	Dim d, classNM
	Dim dtDate
	Dim CalCell

	If lgInputLock Then
		Exit Sub
	End If
	
	Err.Clear
	On Error Resume Next
	Set CalCell = document.all.calendar.tBodies.dayList.rows(i).cells(j)
	
	If Err.number <> 0 Then
		Exit Sub
	End If
	
	d = CalCell.innertext
	classNm = CalCell.className
	
	If Trim(d) <> "" then
		lgDay = CalCell.innertext
		lgCurrCell.className = "Day"
		
		Set lgCurrCell = CalCell
		lgCurrCell.className = "SelDay"	

		lgCurrI = i
		lgCurrJ = j

		If classNM <> "DummyDay" Then '이번 달 
			txtDate.value = UNIFormatDate(CDate(lgYear & "-" & lgMonth & "-" & lgDay))
		Else '이전 달 or 다음 달 
			dtDate = CDate(lgYear & "-" & lgMonth & "-" & "1")
			If CInt(d) > 15 Then '이전 달 
				dtDate = DateAdd("m", -1, dtDate)
			Else '다음 달 
				dtDate = DateAdd("m", 1, dtDate)
			End If

			Call SetDate(Year(dtDate), Month(dtDate), d)
		End If
	End If
	
	Set CalCell = Nothing
End sub

Sub document_onkeydown()
	Dim i, j

	If lgInputLock OR UCase(TypeName(window.event.srcElement)) = "HTMLSELECTELEMENT" Then
		Exit Sub
	End If
	
	Select Case window.event.keyCode
		Case 13 'vbKeyReturn
			Call btnOK_onclick
		Case 27 'vbKeyEscape
			Call btnCancel_onclick
		Case 37 'vbKeyLeft
			If lgCurrJ = 0 Then
				i = lgCurrI - 1
				j = 6
			Else
				i = lgCurrI
				j = lgCurrJ - 1
			End If
			
			Call DayClick(i,j)
		Case 38 'vbKeyUp
			i = lgCurrI - 1
			j = lgCurrJ
			
			Call DayClick(i,j)
		Case 39 'vbKeyRight
			If lgCurrJ = 6 Then
				i = lgCurrI + 1
				j = 0
			Else
				i = lgCurrI
				j = lgCurrJ + 1			
			End If

			Call DayClick(i,j)
		Case 40 'vbKeyDown
			i = lgCurrI + 1
			j = lgCurrJ
		
			Call DayClick(i,j)
		Case 33 'vbKeyPageUp
			Call ChangeMonth(1)
		Case 34 'vbKeyPageDown
			Call ChangeMonth(-1)
	End Select
End Sub

Sub cboYear_onkeyup()
	Call cboYear_onClick
End Sub

Sub cboYear_onClick()
	If CInt(lgYear) <> CInt(cboYear.value) And Not lgInputLock Then
		lgYear = cboYear.value

		Call SetDate(lgYear, lgMonth, lgDay)
	End If
End Sub

Sub cboMonth_onkeyup()
	Call cboMonth_onClick
End Sub

Sub cboMonth_onClick()
	If CInt(lgMonth) <> CInt(cboMonth.value) And Not lgInputLock Then
		lgMonth = cboMonth.value

		Call SetDate(lgYear, lgMonth, lgDay)
	End If
End Sub

Sub ChangeMonth(i)
	Dim dtDate
	
	dtDate = CDate(lgYear & "-" & lgMonth & "-" & lgDay)
	dtDate = DateAdd("m", i, dtDate)
	
	Call SetDate(Year(dtDate), Month(dtDate), Day(dtDate))
End Sub

Function SetDate(y, m, d)
	Dim dtDate
	
	If y<1 OR m<1 OR m>12 OR d<1 OR d>31 Then
		SetDate = False
		Exit Function
	End If

	On Error Resume Next
	Err.Clear

	dtDate = CDate(y & "-" & m & "-" & d)
	
	Do While Err.number <> 0
		If d<27 Then
			SetDate = False
			Exit Function
		End If
		d = CInt(d) - 1
		dtDate = CDate(y & "-" & m & "-" & d)
	Loop
	
	lgYear = Year(dtDate)
	lgMonth = Month(dtDate)
	lgDay = Day(dtDate)	

	'Set Title
	document.all.tbTitle.rows(0).cells(0).innerText = lgYear & ". " & lgMonth

	'Set Date
	txtDate.value = UNIFormatDate(CDate(lgYear & "-" & lgMonth & "-" & lgDay))
	
	'Set Combo	
	cboYear.value = CStr(lgYear)
	cboMonth.value = Right("0" & lgMonth, 2)
	
	lgInputLock = True
	Call SetTimeout("CreateCal", 10)
	
	SetDate = True
End Function

Sub btnOk_onClick()
	Selector = "Y"
	window.returnValue = txtDate.Value
	window.close() 
End Sub

Sub btnCancel_onClick()
	Selector = "N"
	window.returnValue = ""
	window.close()
End Sub

-->
</SCRIPT>

</HEAD>
<BODY bgColor=#E7E7E7 scroll=No>
<TABLE Class="basicTB" CELLSPACING=0>
	<TR>
		<TD>
			<TABLE ID="tbTitle" WIDTH=95% BORDER=0 CELLSPACING=0 ALIGN="center">
				<TR>
					<TD Class="CalTitle" WIDTH=*>&nbsp;</TD>
					<TD WIDTH=10px>
						<SELECT Name="cboYear" tabindex=-1>
							<OPTION Value="<%=Year(Date)%>"><%=Year(Date)%></OPTION>
						</SELECT>
					</TD>
					<TD WIDTH=10px>
						<SELECT Name="cboMonth" tabindex=-1>
							<OPTION Value="<%=Right("0" & Month(Date), 2)%>"><%=Right("0" & Month(Date), 2)%></OPTION>
						</SELECT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
        <TD>
			<TABLE ID="Calendar" WIDTH=95% HEIGHT=100% BORDER=1 CELLSPACING=0 CELLPADDING=0 ALIGN="center">
				<THEAD CLASS="Header">
					<TR>
						<TD WIDTH=<%=100/7 & "%"%>>Sun</TD>
						<TD WIDTH=<%=100/7 & "%"%>>Mon</TD>
						<TD WIDTH=<%=100/7 & "%"%>>Tue</TD>
						<TD WIDTH=<%=100/7 & "%"%>>Wed</TD>
						<TD WIDTH=<%=100/7 & "%"%>>Thu</TD>
						<TD WIDTH=<%=100/7 & "%"%>>Fri</TD>
						<TD WIDTH=<%=100/7 & "%"%>>Sat</TD>
			        </TR>
		    	</THEAD>
				<TBODY ID="dayList">
<%
Dim i, j

For i=0 To 5
%>
			        <TR>
<%
	For j=0 To 6
%>
						<TD  CLASS="Day" ALIGN="Center" onmousedown="DayClick <%=i%>, <%=j%>" ondblclick="DayDblClick <%=i%>, <%=j%>">&nbsp;</TD>
<%
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
	<TR VAlign=center>
		<TD Height=30>
			<TABLE WIDTH=95% HEIGHT=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN="center">
				<TR>
					<TD> <SPAN CLASS="normal">Date</Span> </TD>
					<TD Width=*>
						<INPUT name=txtDate style="TEXT-ALIGN: center;" SIZE=10 ReadOnly Tabindex="-1">
					</TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../image/ok_d.gif" NAME="btnOK" Style="CURSOR: hand" ALT="OK" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../image/cancel_d.gif" NAME="btnCancel" Style="CURSOR: hand" ALT="CANCEL" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>
