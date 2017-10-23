<!-- #Include file="../ESSinc/incServer.asp"  -->
<%
	SYear = Year(Now)
	SMonth = Month(Now)
	SDay = Day(Now)
%>
<HTML>
<HEAD>

<Link Rel="stylesheet" Type="Text/css" HREF="../ESSinc/common.css">
<STYLE TYPE="text/css">
	.CalTitle {font-size:20; font-weight:bold; text-align:center}
	.Header {height:20; font-weight:bold; text-align: center;background-color:linen; color:brown; valign:center}
	.Day {text-align:center; valign:center; background-color:LIGHTCYAN; color:darkblue}
</STYLE>

<TITLE>Calendar</TITLE>

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/Ccm.vbs">      </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/inccommFunc.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../ESSinc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
<!--
Option Explicit


Dim Selector
Dim lgCurrCell
Dim lgCurrI, lgCurrJ
Dim lgYear, lgMonth, lgDay, lgWeekday
Dim lgInputLock
Dim dtTemp

Err.Clear

dtTemp =Trim(window.dialogArguments)

If dtTemp = "" or checkDateFormat(dtTemp, gDateFormat) = false Then
	lgYear  = <%=SYear%>
	lgMonth = <%=SMonth%>
	lgDay   = <%=SDay%>
Else
	Call ExtractDateFrom(dtTemp, gDateFormat, gComDateType,lgYear,lgMonth,lgDay)
	
End If

lgInputLock = False

Sub CreateCal()
	Dim i, j 
	Dim CalRow, CalCol
	Dim CalCell
	Dim dtDate


	dtDate = uniConvDateAtoB(lgYear & "-" & lgMonth & "-" & "01", gServerDateFormat, gClientDateFormat)
	' -- 1일 이전 데이타 클리어 --- 
	For i = (WeekDay(dtDate) - 1) -1 To 0 Step -1
		Set CalCell = document.all.calendar.tBodies.dayList.rows(0).cells(i)

		dtDate = DateAdd("d", -1, dtDate)

		CalCell.innerText = CStr(Day(dtDate))
		CalCell.className = "DummyDay"
	Next
	
	dtDate = uniConvDateAtoB(lgYear & "-" & lgMonth & "-" & "01", gServerDateFormat, gClientDateFormat)

	CalRow = 0
	CalCol = WeekDay(dtDate) - 1
	
	Do While CInt(Month(dtDate)) = CInt(lgMonth)
		Set CalCell = document.all.calendar.tBodies.dayList.rows(CalRow).cells(CalCol)
		
		CalCell.innerText = CStr(Day(dtDate))
		
		If CInt(CalCell.innerText) = CInt(lgDay) Then
			CalCell.className = "pucldrow_tod"
			Set lgCurrCell = CalCell
			
			lgCurrI = CalRow
			lgCurrJ = CalCol
		Else
			If WeekDay(dtDate) = 1 Then
				CalCell.className = "pucldrow_sun"
			ElseIf WeekDay(dtDate) = 7 Then
				CalCell.className = "pucldrow_sat"
			Else 
				CalCell.className = "pucldrow"
			End If
		End If
		
		If WeekDay(dtDate) = 7 Then
			CalRow = CalRow + 1 
			CalCol = 0
		Else
			CalCol = CalCol + 1
		End If
		
		dtDate = DateAdd("d", 1, dtDate)

	Loop
	
	' -- 마지막 날 이후 데이타 클리어 --- %>
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
	Dim i, ii, ii_text
	Dim oOption
	
	For i=lgYear-100 To lgYear+100
		Call SetCombo(cboYear, i, i)
	Next

	cboYear.remove 0
    cboYear.value = CStr(lgYear)
    
	For i=1 To 12
		ii = Right("0" & i, 2)
		ii_text = ii & "월"
		Call SetCombo(cboMonth, ii, ii_text)
	Next

	cboMonth.remove 0
	cboMonth.value = Right("0" & lgMonth, 2)
End Sub

Sub window_onload

	'Call GetGlobalVar()
	Selector = "N"
	Call InitComboBox

	Call SetDate(lgYear, lgMonth, lgDay)

	Call MM_preloadImages("../../CShared/ESSimage/Query.gif","../../CShared/ESSimage/OK.gif","../../CShared/ESSimage/Cancel.gif")

End Sub

Sub window_onunload
	if Selector = "Y" then
		window.returnValue = uniConvYYYYMMDDToDate(gDateFormat,lgYear,lgMonth,lgDay)
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

	Set CalCell = document.all.calendar.tBodies.dayList.rows(i).cells(j)
	
	If Err.number <> 0 Then
		Exit Sub
	End If
	
	d = CalCell.innertext
	classNm = CalCell.className

	If Trim(d) <> "" then

		If lgWeekday = 1 Then
			lgCurrCell.className = "pucldrow_sun"
		ElseIf lgWeekday = 7 Then
			lgCurrCell.className = "pucldrow_sat"
		Else 
			lgCurrCell.className = "pucldrow"
		End If
		
		Set lgCurrCell = CalCell
		lgCurrCell.className = "pucldrow_tod"	

		lgCurrI = i
		lgCurrJ = j
		
		If classNM <> "DummyDay" Then '이번 달 
			lgDay = CalCell.innertext
			dtDate = uniConvYYYYMMDDToDate(gDateFormat,lgYear,lgMonth,lgDay)
			txtDate.value = lgDay & "일" & " " & SetWeekDay(Weekday(dtDate))
		Else '이전 달 or 다음 달 
			dtDate = CDate(lgYear & "-" & lgMonth & "-" & "1")
			If CInt(d) > 15 Then '이전 달 
				dtDate = DateAdd("m", -1, dtDate)
			Else '다음 달 
				dtDate = DateAdd("m", 1, dtDate)
			End If
		End If
	End If

	Call SetDate(Year(dtDate), Month(dtDate), d)
	
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
	lgWeekday = Weekday(dtDate)

	'Set Title
	'document.all.tbTitle.rows(0).cells(0).innerText = lgYear & ". " & lgMonth

	'Set Date
	txtDate.value = lgDay & "일" & " " & SetWeekDay(Weekday(lgWeekDay))
	
	'Set Combo	
	cboYear.value = CStr(lgYear)
	cboMonth.value = Right("0" & lgMonth, 2)
	
	lgInputLock = True
	Call SetTimeout("CreateCal", 10)
	
	SetDate = True
End Function

Sub btnOk_onClick()
	Selector = "Y"
	window.returnValue = uniConvYYYYMMDDToDate(gDateFormat,lgYear,lgMonth,lgDay)
	window.close() 
End Sub

Sub btnCancel_onClick()
	Selector = "N"
	window.returnValue = ""
	window.close()
End Sub

Function SetWeekDay(inWeekDay)

	Dim conv_WeekDay(7) 
	
	conv_WeekDay(0) = "일요일" 
	conv_WeekDay(1) = "월요일" 
	conv_WeekDay(2) = "화요일" 
	conv_WeekDay(3) = "수요일" 
	conv_WeekDay(4) = "목요일"
	conv_WeekDay(5) = "금요일"
	conv_WeekDay(6) = "토요일"
	
	SetWeekDay = conv_WeekDay(inWeekDay-1)
End Function

-->
</SCRIPT>

</HEAD>
<body>
<table width="196" border="0" cellspacing="1" cellpadding="0" bgcolor="#949494" align=center valign=middle>
  <tr> 
    <td align="center" background="../../CShared/ESSimage/bg_calendar_tp.gif">
	  <table ID="tbTitle" width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr> 
          <td><img src="../../CShared/ESSimage/im_calendar_lt.gif" border="0"></td>
          <td width="7"><img border="0" src="trans.gif" width="7" height="1"></td>
          <td class="ftgray" width=30 ><font color="#000000">
			<SELECT CLASS="form01" Name="cboYear" tabindex=-1><OPTION Value="<%=Year(Date)%>"><%=Year(Date)%></OPTION></SELECT></font>
          </td>
          <td width="7"><img border="0" src="trans.gif" width="7" height="1"></td>
          <td><img src="../../CShared/ESSimage/im_calendar_rt.gif" width="41" height="28" border="0"></td>
          <td align="right"><img src="../../CShared/ESSimage/ic_calendar_close.gif" NAME="btnCancel" ALT="CANCEL" onclick="VBSCRIPT:CALL btnCancel_onclick()" width="15" height="28" border="0" style="cursor: hand" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/ESSimage/ic_calendar_close.gif',1)"></td>
          <td width="7"><img border="0" src="trans.gif" width="7" height="1"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td bgcolor="#ffffff" align="center"> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center"> <table border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td><img src="../../CShared/ESSimage/ic_sb_arrow_prev.gif" onclick="vbscript:call ChangeMonth(-1)" Style="CURSOR: hand" width="6" height="13" border="0"></td>
                <td width="7"><img border="0" src="trans.gif" width="5" height="1"></td>
                <td class="ftblack">
					<SELECT CLASS="form01" Name="cboMonth" tabindex=-1>
					    <OPTION Value="<%=Right("0" & Month(Date), 2)%>"><%=Right("0" & Month(Date), 2)%></OPTION>
					</SELECT>
                </td>
                <td width="7"><img border="0" src="trans.gif" width="5" height="1"></td>
                <td><img src="../../CShared/ESSimage/ic_sb_arrow_next.gif" onclick="vbscript:call ChangeMonth(1)" Style="CURSOR: hand" width="6" height="13" border="0"></td>
                <td width="7"><img border="0" src="trans.gif" width="7" height="1"></td>
                <td>
					<INPUT class="ftsblue" name=txtDate SIZE=25 ReadOnly Tabindex="-1"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td align="center"> 
            <table ID="Calendar" border="0" cellpadding="0" cellspacing="1" bgcolor="#EEEBE2">
			  <THEAD CLASS="Header">
				<tr> 
				  <td width="20" class="pucldhead">일</td>
				  <td width="20" class="pucldhead">월</td>
				  <td width="20" class="pucldhead">화</td>
				  <td width="20" class="pucldhead">수</td>
				  <td width="20" class="pucldhead">목</td>
				  <td width="20" class="pucldhead">금</td>
				  <td width="20" class="pucldhead">토</td>
				</tr>
		      </THEAD>
			  <TBODY ID="dayList">
			<%
				Dim i, j, class_nm

				For i=0 To 5
			%>
			    <TR>
			<%
					For j=0 To 6					
			%>
					<TD ALIGN="Center" bgcolor="#E3E3E3" Style="CURSOR: hand" onmousedown="DayClick <%=i%>, <%=j%>" ondblclick="DayDblClick <%=i%>, <%=j%>">&nbsp;</TD>
			<%
					Next
			%>
				</TR>
			<%
				Next
			%>
			  </TBODY>
            </table>
          </td>
        </tr>
	</td>
  </tr>
</TABLE>
</BODY>
</HTML>
