<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Management
'*  3. Program ID           : a7114mb1
'*  4. Program Name         : �����󰢰���ݿ� 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             AS0071
'                             
'*  7. Modified date(First) : 2000/3/21
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	On Error Resume Next														'��: 
	           																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

	Call HideStatusWnd
	Call LoadBasisGlobalInf()

	strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

	Select Case strMode
	Case CStr(UID_M0002)		
	     Call SubBizBatch()
		
	End Select
	
	Response.End
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()
           																	'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim strWKfg
Dim iPAAG070
Dim I1_com_asst_var
Dim I2_a_asset_depr_master
Dim I3_ief_supplied

    Err.Clear                                                               '��: Protect system from crashing    
	On Error Resume Next														'��: 

    'If  Request("txtWKyymm") = "" Then   
     '   Call ServerMesgBox("�۾����س���� �����Ͻʽÿ�.", vbInformation, I_MKSCRIPT)              
	'	Response.End 
	'End If
	
	Set iPAAG070 = Server.CreateObject("PAAG070.cAExecRefltSvr")
		
    '-------------------------------------------
    'Com action result check area(OS,internal)
    '-------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG070 = Nothing
       Exit Sub
    End If    
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I3_ief_supplied	 =  Request("txtRadio")
    I1_com_asst_var   =  Request("txtWKyymm")
    I2_a_asset_depr_master = gUsrId  
         
	Call iPAAG070.AS0071_EXECUTE_REFLT_SVR(gStrGlobalCollection ,I1_com_asst_var, I2_a_asset_depr_master,I3_ief_supplied )

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG070 = Nothing
       Exit Sub
    End If    

	Response.Write " <Script Language=vbscript>	                           " & vbCr
	Response.Write "    parent.fnButtonExecOk                              " & vbCr
    Response.Write " </Script>                                             " & vbCr
    Response.End    

    Set iPAAG025 = Nothing														    '��: Unload Comproxy
end sub 
%>