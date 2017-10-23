<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	

'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  p1102mb1.asp
'*  4. Program Name         :  Mfg Calendar 조회 
'*  5. Program Desc         :
'*  6. Component List		: PP1G102.cPListMfgCalenSvr
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2002/05/09
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 


    Const P104_I1_temp_month = 0
    Const P104_I1_temp_year = 1
    
    '[CONVERSION INFORMATION]  EXPORTS View 상수 
    '[CONVERSION INFORMATION]  View Name : export_next p_mfg_calendar
    Const P104_E1_dt = 0
    '[CONVERSION INFORMATION]  View Name : export_item p_mfg_calendar_type
    Const P104_EG1_E1_p_mfg_calendar_type_cal_type = 0
    Const P104_EG1_E1_p_mfg_calendar_type_cal_type_nm = 1
    '[CONVERSION INFORMATION]  View Name : export_item p_mfg_calendar
    Const P104_EG1_E2_p_mfg_calendar_dt = 2
    Const P104_EG1_E2_p_mfg_calendar_julian_day = 3
    Const P104_EG1_E2_p_mfg_calendar_day_of_the_week = 4
    Const P104_EG1_E2_p_mfg_calendar_work_type = 5
    Const P104_EG1_E2_p_mfg_calendar_yr_accum_work_day = 6
    Const P104_EG1_E2_p_mfg_calendar_tot_accum_work_day = 7
    Const P104_EG1_E2_p_mfg_calendar_lot_perd_no = 8
    Const P104_EG1_E2_p_mfg_calendar_remark = 9

On Error Resume Next														'☜: 
    Err.Clear                                                               '☜: Protect system from crashing
    
	Dim pPP1G102
	Dim I1_prod_work_set
	Dim I2_p_calendar_type_cal_type
	Dim pvSheetMaxRowsD
	Dim E1_p_mfg_calendar
	Dim EG1_export_group
	
	Dim dtDate
	Dim startIndex
	Dim i
	Dim CurDate
	Dim iErrorPosition
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call LoadBasisGlobalInf() 
	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

	CurDate = GetSvrDate

    
    If Request("txtYear") = "" Or Request("txtMonth") = "" Then				'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	End If
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    Redim I1_prod_work_set(P104_I1_temp_year)
    
    I2_p_calendar_type_cal_type = Trim(Request("txtClnrType"))
	I1_prod_work_set(P104_I1_temp_month) = Trim(Request("txtMonth"))
	I1_prod_work_set(P104_I1_temp_year)  = Trim(Request("txtYear"))
	
    '-----------------------
    'Com action area
    '-----------------------
    Set pPP1G102 = Server.CreateObject("PP1G102.cPListMfgCalenSvr")

    If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
	
	Call pPP1G102.P_LIST_MFG_CALENDAR  (gStrGlobalCollection, I1_prod_work_set, I2_p_calendar_type_cal_type, _
			E1_p_mfg_calendar, EG1_export_group)
			
	If CheckSYSTEMError2(Err, True, iErrorPosition & "일", "", "", "", "") = True Then
		Set pPP1G102 = Nothing															'☜: Unload Component
		Response.End
	End If
	
	Set pPP1G102 = Nothing


	'-----------------------
	'Result data display area
	'----------------------- 
    dtDate = EG1_export_group(0, P104_EG1_E2_p_mfg_calendar_dt)
      
    
    Call ExtractDateFrom(dtDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)
    startIndex = WeekDay(dtDate) -1 
  
%>
<Script Language=VBScript>
	parent.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(EG1_export_group(0, P104_EG1_E1_p_mfg_calendar_type_cal_type_nm))%>" 
	Parent.frm1.txtYear.value = "<%=strYear%>"
	Parent.frm1.txtMonth.value = "<%=strMonth%>"
	
	parent.frm1.cboYear.value =	 "<%=strYear%>"
	parent.frm1.cboMonth.value = "<%=strMonth%>"
	
	Parent.lgStartIndex = "<%=startIndex%>"

<%
	dtDate = UNIDateAdd("m",1,dtDate,gServerDateFormat)
	dtDate = UNIDateAdd("d", -1,dtDate,gServerDateFormat)
	Call ExtractDateFrom(dtDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)

    dtDate = UNIDateAdd("d", -1,EG1_export_group(0,P104_EG1_E2_p_mfg_calendar_dt),gServerDateFormat)
    Call ExtractDateFrom(dtDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)
%>
	Parent.lgLastDay = "<%=strDay%>"

	Dim CalCol


	<%' -- 1일 이전 데이타 클리어 --- %>
	For CalCol = <%=startIndex-1%> To 0 Step -1
		Parent.frm1.txtDate(CalCol).value = CStr(<%=strDay%> + CalCol - <%=startIndex-1%>)
		Parent.frm1.txtDate(CalCol).className = "DummyDay"
		Parent.frm1.txtDate(CalCol).disabled = True
		
		Parent.frm1.txtHoli(CalCol).value = ""
		Parent.frm1.txtHoli(CalCol).disabled = True

		Parent.frm1.txtDesc(CalCol).value = ""
		Parent.frm1.txtDesc(CalCol).disabled = True
		Parent.frm1.txtDesc(CalCol).title = ""
	Next
<%
    For i = 1 To ubound(EG1_export_group, 1) + 1
    	If EG1_export_group(i-1, P104_EG1_E2_p_mfg_calendar_work_type) = "0" Then
%>
			Parent.frm1.txtDate(<%=startIndex%>).style.color = "red"
<%
		Else

			If EG1_export_group(i-1, P104_EG1_E2_p_mfg_calendar_work_type) = "1" Then
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
		Parent.frm1.txtDate(<%=startIndex%>).value = "<%=i%>"
		Parent.frm1.txtDate(<%=startIndex%>).className = "Day"
		Parent.frm1.txtDate(<%=startIndex%>).disabled = False
	
		Parent.frm1.txtHoli(<%=startIndex%>).value = "<%=EG1_export_group(i-1, P104_EG1_E2_p_mfg_calendar_work_type)%>"
		
		Parent.frm1.txtHoli(<%=startIndex%>).disabled = False
	
		Parent.frm1.txtDesc(<%=startIndex%>).value = "<%=ConvSPChars(EG1_export_group(i-1, P104_EG1_E2_p_mfg_calendar_remark))%>"
		<%	
			If CDate(EG1_export_group(i-1,P104_EG1_E2_p_mfg_calendar_dt)) < CDate(CurDate) Then
		%>
				Parent.frm1.txtDesc(<%=startIndex%>).disabled = True
		<%
			Else
		%>
				Parent.frm1.txtDesc(<%=startIndex%>).disabled = False
		<%
			End If
		%>
			Parent.frm1.txtDesc(<%=startIndex%>).title = "<%=ConvSPChars(EG1_export_group(i-1, P104_EG1_E2_p_mfg_calendar_remark))%>"	
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

	Parent.lgNextNo = ""									' 다음 키 값 넘겨줌 
	Parent.lgPrevNo = ""									' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 
		
	Parent.DbQueryOk										'☜: 조회가 성공 
</Script>
<%
	Response.End											'☜: Process End
%>
