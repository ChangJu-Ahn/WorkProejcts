<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  p1102mb3.asp
'*  4. Program Name         :  Mfg Calendar ���� 
'*  5. Program Desc         :
'*  6. Component List		: PP1G103.cPMngMfgCalen
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next																		'��: 

'Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
'Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 


    '[CONVERSION INFORMATION]  IMPORTS View ��� 
	'[CONVERSION INFORMATION]  View Name : import prod_work_set
     Const P110_I1_temp_month = 0    
     Const P110_I1_temp_year = 1
     													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd																			'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Call LoadBasisGlobalInf() 


Dim pPP1G103
Dim IG1_group_of_saturdau_count
Dim I1_prod_work_set
Dim I2_p_mfg_calendar_type_cal_type
Dim IG2_group_of_day 
Dim iCommandSent
Dim IG3_import_group_of_date

Dim strMode																					'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim dtDate
Dim i
Dim lgIntFlgMode

	strMode = Request("txtMode")															'�� : ���� ���¸� ���� 
									
    Err.Clear																				'��: Protect system from crashing
	
    If Request("txtFlgMode") = "" Then														'��: ������ ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	ElseIf Request("txtClnrType") = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)              
		Response.End                  		
	End If

	lgIntFlgMode = CInt(Request("txtFlgMode"))												'��: ����� Create/Update �Ǻ� 

    '-----------------------
    'Data manipulate area
    '-----------------------
    Redim I1_prod_work_set(P110_I1_temp_year)
	Redim IG3_import_group_of_date(cint(Request("txtHoli").count), 2)
	
		I2_p_mfg_calendar_type_cal_type = Request("txtClnrType")
		I1_prod_work_set(P110_I1_temp_year) = Request("txtYear")
		I1_prod_work_set(P110_I1_temp_month) = Right("0" & Request("txtMonth"),2)
		
    For i = 0 To Request("txtHoli").count 
    
        IG3_import_group_of_date(i, 0)	= Request("txtDesc")(i)
        
		dtDate = UNIConvDate(UNIConvYYYYMMDDToDate(gDateFormat, Request("txtYear"), Request("txtMonth"), CStr(i)))
        IG3_import_group_of_date(i, 1) = dtDate
        IG3_import_group_of_date(i, 2) = Request("txtHoli")(i)        
        
    Next

    If lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If
    
    Set pPP1G103 = Server.CreateObject("PP1G103.cPMngMfgCalen")   
    
    If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
	
														
	call pPP1G103.P_MANAGE_MFG_CALENDAR (gStrGlobalCollection, IG1_group_of_saturdau_count, I1_prod_work_set, _
				I2_p_mfg_calendar_type_cal_type, IG2_group_of_day, IG3_import_group_of_date, iCommandSent)
				
	If CheckSYSTEMError(Err,True) = True Then
		Set pPP1G103 = Nothing												'��: ComProxy Unload
		Response.End
	End If
	
	Set pPP1G103 = Nothing												'��: ComProxy Unload
	

	'-----------------------
	'Result data display area
	'----------------------- 
%>

<Script Language=vbscript>

	Parent.DbSaveOk
</Script>
<%					
	
	Response.End																				'��: Process End
	
%>