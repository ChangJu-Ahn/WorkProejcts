<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Calendar수정)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B15012UpdateCalendar
'                             +B15018ListCalendar
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/09/19
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

'On Error Resume Next														'☜: 

Dim pB15018
Dim pB15012

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim dtDate
Dim startIndex
Dim lastDay
Dim i
Dim lgIntFlgMode


Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    If Request("txtYear") = "" Or Request("txtMonth") = "" Then				'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If


    Dim strYYYYMM
    Dim I1_apprl_yrmnth

    Const A583_EG1_E1_calendar_dt = 0
    Const A583_EG1_E1_day_of_week = 1
    Const A583_EG1_E1_hol_type = 2
    Const A583_EG1_E1_remark = 3
    Const A583_EG1_E1_gl_fg = 4
    Const A583_EG1_E1_temp_gl_fg = 5

	Dim ObjPB4G031
	Dim Export_Array
    
    strYYYYMM = Right("0000" & Request("txtYear"), 4)
    strYYYYMM = strYYYYMM & Right("00" & Request("txtMonth"), 2)

    I1_apprl_yrmnth = strYYYYMM

    Set ObjPB4G031 = server.CreateObject ("PB4G031.cBListCalendar")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB4G031.B_LIST_CALENDAR (gStrGlobalCollection,I1_apprl_yrmnth)
    Set ObjPB4G031 = nothing

    If CheckSYSTEMError(Err,True) = True Then                       
        Response.End 
    End If
    on error goto 0

	'-----------------------
	'Result data display area
	'----------------------- 
    'dtDate = CDate(pB15018.ExportItemBCalendarCalendarDt(1))
    dtDate = CDate(Trim(Export_Array(0,A583_EG1_E1_calendar_dt)))
    startIndex = WeekDay(dtDate) - 1
    
%>
<Script Language=VBScript>
	Parent.frm1.hYear.value = "<%=Year(dtDate)%>"
	Parent.frm1.hMonth.value = "<%=Month(dtDate)%>"
	Parent.document.all.tbTitle.Rows(0).Cells(0).innerText = "<%=Year(dtDate)%>" & ". " & "<%=Month(dtDate)%>"
	parent.frm1.txtYymm.text = "<%=UNIMonthClientFormat(dtDate)%>"
	
	Parent.lgStartIndex = <%=startIndex%>
<%
	dtDate = DateAdd("m", 1, dtDate)
	dtDate = DateAdd("d", -1, dtDate)
	lastDay = Day(dtDate)
	
	'지난달 Display를 위해서....
    'dtDate = CDate(pB15018.ExportItemBCalendarCalendarDt(1))
     dtDate = CDate(Trim(Export_Array(0,A583_EG1_E1_calendar_dt)))
%>
	Parent.lgLastDay = <%=lastDay%>

	Dim CalCol

	<%' -- 1일 이전 데이타 클리어 --- %>
	For CalCol = <%=startIndex-1%> To 0 Step -1
		Parent.frm1.txtDate(CalCol).value = CStr(<%=Day(DateAdd("d", -1, dtDate))%> + CalCol - <%=startIndex-1%>)
		Parent.frm1.txtDate(CalCol).className = "DummyDay"
		Parent.frm1.txtDate(CalCol).disabled = True
		
		Parent.frm1.txtHoli(CalCol).value = ""
		Parent.frm1.txtHoli(CalCol).disabled = True

		Parent.frm1.txtDesc(CalCol).value = ""
		Parent.frm1.txtDesc(CalCol).disabled = True
		Parent.frm1.txtDesc(CalCol).title = ""
	Next
	
<%
    GroupCount = Ubound(Export_Array,1)
	For i = 0 To GroupCount
		If Trim(Export_Array(i,A583_EG1_E1_hol_type)) = "H" Then
%>
	Parent.frm1.txtDate(<%=startIndex%>).style.color = "red"
<%
		Else
			If (startIndex + 1) Mod 7 = 0 Then
%>
	Parent.frm1.txtDate(<%=startIndex%>).style.color = "blue"
<%	
			Else
%>
	Parent.frm1.txtDate(<%=startIndex%>).style.color = "black"
<%
			End If
		End If
%>
	Parent.frm1.txtDate(<%=startIndex%>).value = "<%=i+1%>"
	Parent.frm1.txtDesc(<%=startIndex%>).alt = "<%=i+1%>" & "일의 사유"
	Parent.frm1.txtDate(<%=startIndex%>).className = "Day"
	Parent.frm1.txtDate(<%=startIndex%>).disabled = False
	
	Parent.frm1.txtHoli(<%=startIndex%>).value = "<%=ConvSPChars(Trim(Export_Array(i,A583_EG1_E1_hol_type)))%>"
	Parent.frm1.txtHoli(<%=startIndex%>).disabled = False
	
	Parent.frm1.txtDesc(<%=startIndex%>).value = "<%=ConvSPChars(Trim(Export_Array(i,A583_EG1_E1_remark)))%>"
	Parent.frm1.txtDesc(<%=startIndex%>).disabled = False
	Parent.frm1.txtDesc(<%=startIndex%>).title = "<%=ConvSPChars(Trim(Export_Array(i,A583_EG1_E1_remark)))%>"	
<%
		startIndex = startIndex + 1
	Next	
%>
	For CalCol = <%=startIndex%> to 41
		Parent.frm1.txtDate(CalCol).value = CStr(CalCol - <%=startIndex-1%>)
		Parent.frm1.txtDate(CalCol).className = "DummyDay"
		Parent.frm1.txtDate(CalCol).disabled = True

		Parent.frm1.txtHoli(CalCol).value = ""
		Parent.frm1.txtHoli(CalCol).disabled = True

		Parent.frm1.txtDesc(CalCol).value = ""
		Parent.frm1.txtDesc(CalCol).disabled = True
		Parent.frm1.txtDesc(CalCol).title = ""
	Next

	Parent.lgNextNo = ""		' 다음 키 값 넘겨줌 
	Parent.lgPrevNo = ""		' 이전 키 값 넘겨줌 
		
	Parent.DbQueryOk																'☜: 조회가 성공 
</Script>
<%

	Response.End																				'☜: Process End

Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
	Dim strVal
    Dim Obj2PB4G031
	strVal = ""
    
    If Request("txtFlgMode") = "" Then											'⊙: 저장을 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("TXTFLGMODE 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

    '-----------------------
    'Data manipulate area
    '-----------------------
    For i = 1 To Request("txtHoli").count
		dtDate = CDate(Request("hYear") & "-" & Request("hMonth") & "-" & i)
        strVal = strVal & Trim(dtDate) & gColSep
        strVal = strVal & Trim(Request("txtHoli")(i)) & gColSep
        strVal = strVal & Trim(Request("txtDesc")(i)) & gRowSep
    Next
    
    
    
    Set Obj2PB4G031 = server.CreateObject ("PB4G031.cBUptCalendar")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB4G031.B_UPDATE_CALENDAR (gStrGlobalCollection,Trim(strVal))
    Set Obj2PB4G031 = nothing

    If CheckSYSTEMError(Err,True) = True Then                       
		Response.End 
    End If
    on error goto 0                                                             



%>
<Script Language="VBScript">
	Call parent.DbSaveOk()
</Script>
<%					
    Set pB15012 = Nothing                                                   '☜: Unload Comproxy

	Response.End																				'☜: Process End
	
End Select
%>