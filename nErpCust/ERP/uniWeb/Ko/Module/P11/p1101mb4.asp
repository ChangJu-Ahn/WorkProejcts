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

'Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
'Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
												'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

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
	
	Call HideStatusWnd																			'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
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
			With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
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
			With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
				.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(E1_p_mfg_calendar_type(P111_E1_cal_type_nm))%>"
				.DbExecute
			End With
		</Script>
<%	
									
	Else
%>					
		<Script Language=vbscript>
			With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
				.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(E1_p_mfg_calendar_type(P111_E1_cal_type_nm))%>"
				.LotPerdLookUpOk
			End With
		</Script>
<%		
	End If
	Response.End
%>