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
'*  3. Program ID           : a7113mb1(������ ��ǥó��)
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/3/21
'*  8. Modified date(Last)  : 2000/10/26
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

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

%>
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	On Error Resume Next														'��: 
	           																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
	Dim str_FrDt
	Dim str_ToDt

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
  '********************************************************  
  '                        Execution
  '********************************************************  
	Err.Clear
	On Error Resume Next														'��: 
 
    Dim I1_b_biz_area_cd 
    Dim I2_b_acct_dept 
    Dim I3_USR_ID 
    Dim I4_ief_supplied 
	Dim I1_conf_fg
	Dim iPACF310
    
    Const A514_I2_org_change_id = 0    '[CONVERSION INFORMATION]  View Name : import b_acct_dept
    Const A514_I2_dept_cd = 1
        
    Dim I5_job_date
    Dim I5_gl_date

    
    Const A514_E1_dept_cd = 0    '[CONVERSION INFORMATION]  View Name : export b_acct_dept
    Const A514_E1_dept_nm = 1

    Redim I2_b_acct_dept(A514_I2_dept_cd)



	' -- ���Ѱ����߰� 
	Const I7_a_data_auth_data_BizAreaCd = 0
	Const I7_a_data_auth_data_internal_cd = 1
	Const I7_a_data_auth_data_sub_internal_cd = 2
	Const I7_a_data_auth_data_auth_usr_id = 3

	Dim I7_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I7_a_data_auth(3)
	I7_a_data_auth(I7_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I7_a_data_auth(I7_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I7_a_data_auth(I7_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I7_a_data_auth(I7_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

    Err.Clear                                                               '��: Protect system from crashing    
    If  Request("txtGLdt") = "" Then   
        Call ServerMesgBox("117523", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If		
'Call ServerMesgBox("a100", vbInformation, I_MKSCRIPT)   
    '-------------------------------------------
    'Com action result check area(OS,internal)
    '-------------------------------------------
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPACF310 = Nothing
       Exit Sub
    End If    
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ jsk
	gChangeOrgId = Request("txtOrgChangeId")		'GetGlobalInf("gChangeOrgId")
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ jsk
	
	I5_job_date				= Request("txtFrdt")
	I5_gl_date				= Request("txtGldt")
	I1_b_biz_area_cd		= Request("txtBizAreaCd")
	I3_USR_ID  				= gUsrId
		
	I2_b_acct_dept(A514_I2_org_change_id)	= gChangeOrgId
	I2_b_acct_dept(A514_I2_dept_cd)	  		= Trim(Request("txtDeptCd"))	
	I1_conf_fg       						= Request("txtRadio2")'�۾�����  
	I4_ief_supplied						  	= Request("txtRadio") '�۾�����
'Call ServerMesgBox("a200", vbInformation, I_MKSCRIPT)   

	'-----------------------
	'Com action area
	'-----------------------       
'Call ServerMesgBox(I1_conf_fg, vbInformation, I_MKSCRIPT) 	
'Call ServerMesgBox(I4_ief_supplied, vbInformation, I_MKSCRIPT) 
'Call ServerMesgBox(I5_job_date, vbInformation, I_MKSCRIPT) 
'Call ServerMesgBox(I5_gl_date, vbInformation, I_MKSCRIPT) 

	If I1_conf_fg = "U" Then
'Call ServerMesgBox("a210", vbInformation, I_MKSCRIPT) 		
		Set iPACF310 = Server.CreateObject("PACF310_KO441.cAJobBatchWithSpSvr")				
'		Call iPACF310.A_JOB_BATCH_WITH_SP_SVR( gStrGloBalCollection , I1_conf_fg, I4_ief_supplied, I5_job_date, _
' 											   I1_b_biz_area_cd ,  I5_gl_date, I2_b_acct_dept)
		Call iPACF310.A_JOB_BATCH_WITH_SP_SVR( gStrGloBalCollection , I1_conf_fg, I4_ief_supplied, I5_job_date)
'Call ServerMesgBox("a220", vbInformation, I_MKSCRIPT) 											   
	Else
		Set iPACF310 = Server.CreateObject("PACF310_KO441.cAJobToGLSvr")				
		Call iPACF310.A_ToGLSvr( gStrGloBalCollection, I4_ief_supplied , I1_b_biz_area_cd ,  I2_b_acct_dept , _
		                I5_job_date, I5_gl_date)
	End if
'Call ServerMesgBox("a300", vbInformation, I_MKSCRIPT)   	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
    If CheckSYSTEMError(Err, True) = True Then					
       Set iPACF310 = Nothing
       Exit Sub
    End If        
    
	Response.Write " <Script Language=vbscript>	                           " & vbCr
	Response.Write "    With parent                                        " & vbCr
	'Response.Write "	     .frm1.txtDeptNm.value = """   & E1_b_acct_dept(A514_E1_dept_nm)  & """" & vbCr
	Response.Write "		 .lgAnswer = ""Success""						       " & vbCr
	Response.Write "		 .fnButtonExecOk()	                           " & vbCr
    Response.Write " End With                                              " & vbCr
    Response.Write " </Script>                                             " & vbCr
    Response.End    

    Set iPACF310 = Nothing														    '��: Unload Comproxy

	Response.End																		'��: Process End
end sub
%>
