<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1102mb4.asp
'*  4. Program Name         : Look Up Mfg. Calendar
'*  5. Program Desc         :
'*  6. Comproxy List        : +PP1G101.P_LOOK_UP_MFG_CALENDAR.P_LOOK_UP_MFG_CALENDAR
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2000/04/17
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

'Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
'Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
	
	'[CONVERSION INFORMATION]  View Name : export p_mfg_calendar_type
     Const P102_E1_cal_type = 0    
     Const P102_E1_cal_type_nm = 1
     
     '[CONVERSION INFORMATION]  View Name : export p_mfg_calendar
     Const P102_E2_dt = 0    
     Const P102_E2_julian_day = 1
     Const P102_E2_day_of_the_week = 2
     Const P102_E2_work_type = 3
     Const P102_E2_yr_accum_work_day = 4
     Const P102_E2_tot_accum_work_day = 5
     Const P102_E2_remark = 6
     Const P102_E2_lot_perd_no = 7
     
On Error Resume Next														'��: 

	Dim pPP1G101
	Dim I1_p_mfg_calendar_type
	Dim I2_p_mfg_calendar
	Dim pvCommandSent
	Dim E1_p_mfg_calendar_type
	Dim E2_p_mfg_calendar
	Dim E2_p_lot_period_exit
	Dim I2_p_mfg_calendar_dt
	Dim iCommandSent
	
	Call LoadBasisGlobalInf() 

	I2_p_mfg_calendar_dt = Trim(Request("txtYear")) & "-01-01"
	I1_p_mfg_calendar_type = Trim(Request("txtClnrType"))
	iCommandSent = "LIST"
	Set pPP1G101 = server.CreateObject("PP1G101.cPLkUpMfgCalenSvr")
	
    If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
	
	call pPP1G101.P_LOOK_UP_MFG_CALENDAR (gStrGlobalCollection, I1_p_mfg_calendar_type, I2_p_mfg_calendar_dt, _
				iCommandSent, E1_p_mfg_calendar_type, E2_p_mfg_calendar, E2_p_lot_period_exit)
				
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP1G101 = Nothing												'��: ComProxy Unload
		
%>		
		<Script Language=vbscript>
			With parent														'��: ȭ�� ó�� ASP �� ��Ī�� 
				.ClnrNO
			End With
		</Script>

<%		Response.End 
	End If
	
	Set pPP1G101 = Nothing												'��: ComProxy Unload
	
	If E2_p_lot_period_exit="N" Then
%>
		<Script Language=vbscript>
			With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
				.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(E1_p_mfg_calendar_type(P102_E1_cal_type_nm))%>"
				.DbExecute
			End With
		</Script>
<%	
									
	Else
%>

		<Script Language=vbscript>
			With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
				.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(E1_p_mfg_calendar_type(P102_E1_cal_type_nm))%>"
				.ClnrLookUpOk
			End With
		</Script>
<%
	End If		
	Response.End 
	
%>

