<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1101mb4.asp
'*  4. Program Name         : Look Up Lot Period
'*  5. Program Desc         :
'*  6. Component List       : +PP1G104.cPLkUpLotPeriodSvr.P_LOOK_UP_LOT_PERIOD
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2000/04/17
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************

'Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
'Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
												'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														 
	Err.Clear
	
    Const P111_E1_cal_type = 0
    Const P111_E1_cal_type_nm = 1
    
	Dim pPP1G104 
	Dim I1_prod_work_set_temp_timestamp
	Dim I2_p_mfg_calendar_type_cal_type
	Dim iCommandSent
	Dim E1_p_mfg_calendar_type
	Dim E2_p_lot_period
	Dim E2_p_lot_period_exit
	
	Call HideStatusWnd																			'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Call LoadBasisGlobalInf() 
	
	I1_prod_work_set_temp_timestamp  = Trim(Request("txtYear")) & "-01-01"
	I2_p_mfg_calendar_type_cal_type  = Trim(Request("txtClnrType"))
	iCommandSent = "LIST"

	Set pPP1G104 = Server.CreateObject("PP1G104.cPLkUpLotPeriodSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
	
	call pPP1G104.P_LOOK_UP_LOT_PERIOD (gStrGlobalCollection, I1_prod_work_set_temp_timestamp, _
		I2_p_mfg_calendar_type_cal_type, iCommandSent, E1_p_mfg_calendar_type, E2_p_lot_period, E2_p_lot_period_exit)
	
	If CheckSYSTEMError(Err, True) = True Then
		Set pPP1G104 = Nothing	
%>
		<Script Language=vbscript>
			With parent																		'☜: 화면 처리 ASP 를 지칭함 
				.LotPerdNo
			End With
		</Script>
<%	
		Response.End
	End If
	
	Set pPP1G104 = Nothing
	
	If E2_p_lot_period_exit="N" Then
%>
		<Script Language=vbscript>
			With parent																		'☜: 화면 처리 ASP 를 지칭함 
				.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(E1_p_mfg_calendar_type(P111_E1_cal_type_nm))%>"
				.DbExecute
			End With
		</Script>
<%	
									
	Else
%>					
		<Script Language=vbscript>
			With parent																		'☜: 화면 처리 ASP 를 지칭함 
				.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(E1_p_mfg_calendar_type(P111_E1_cal_type_nm))%>"
				.LotPerdLookUpOk
			End With
		</Script>
<%		
	End If
	Response.End
%>