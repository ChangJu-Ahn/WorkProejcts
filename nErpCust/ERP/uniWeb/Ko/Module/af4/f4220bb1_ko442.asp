<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

On Error Resume Next														'��: 
Dim lgOpModeCRUD
Err.Clear  
                                                                      '��: Clear Error status
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","QB")

	Call HideStatusWnd

lgOpModeCRUD = Request("txtMode")												'�� : ���� ���¸� ���� 

Select Case lgOpModeCRUD
     Case CStr(UID_M0002)       
          Call SubBizSave()
End Select

Response.End
'============================================================================================================
' Name : SubBizSave
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear                                                                            '��: Clear Error status

	Dim PAFG400EXE                    				                            '�� : �Է�/������ ComProxy Dll ��� ����(as0031
	Dim EG1_export_group
    Const E1_f_ln_info_loan_nm = 0
    Const E1_ief_supplied_count = 1
    Const E1_f_ln_info_biz_area_nm_from = 2
    Const E1_f_ln_info_biz_area_nm_to = 3    
    
	Dim I2_f_ln_info_biz_area
    Const A755_I2_f_ln_info_from_biz = 0
    Const A755_I2_f_ln_info_to_biz = 1	
	
	Redim I2_f_ln_info_biz_area(1)
	I2_f_ln_info_biz_area(A755_I2_f_ln_info_from_biz) = UCase(Trim(Request("txtBizAreaCd")))
	I2_f_ln_info_biz_area(A755_I2_f_ln_info_to_biz) = UCase(Trim(Request("txtBizAreaCd1")))
	
    Set PAFG400EXE = Server.CreateObject("PAFG400_KO442.cFMngLnPlnSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If

    Call PAFG400EXE.F_MANAGE_LN_PLAN_SVR(gStrGloBalCollection, Request("txtLoanNo"), UniConvDate(Request("txtDateFr")), _
								UniConvDate(Request("txtDateTo")), EG1_export_group,I2_f_ln_info_biz_area)

    If CheckSYSTEMError(Err, True) = True Then
		Set PAFG400EXE = Nothing
		Response.Write "<Script Language=vbscript>	" & vbCr
		Response.Write " With parent.frm1		    " & vbCr       
		Response.Write " .txtBizAreaNm.value  = """ & ConvSPChars(EG1_export_group(E1_f_ln_info_biz_area_nm_from)) & """" & vbCr
		Response.Write " .txtBizAreaNm1.value = """ & ConvSPChars(EG1_export_group(E1_f_ln_info_biz_area_nm_to))   & """" & vbCr       
		Response.Write " End with				" & vbcr
		Response.Write "</Script>               " & vbCr		
		Exit Sub
    End If    

    Set PAFG400EXE = Nothing

	If IsEmpty(EG1_export_group) = False Then
	
		Response.Write "<Script Language=vbscript>	" & vbCr
		Response.Write " With parent.frm1		    " & vbCr
		Response.Write " .txtLoanNm.value     = """ & ConvSPChars(EG1_export_group(E1_f_ln_info_loan_nm))          & """" & vbCr
		Response.Write " .txtCount.value      = """ & ConvSPChars(EG1_export_group(E1_ief_supplied_count))         & """" & vbCr
		Response.Write " .txtBizAreaNm.value  = """ & ConvSPChars(EG1_export_group(E1_f_ln_info_biz_area_nm_from)) & """" & vbCr
		Response.Write " .txtBizAreaNm1.value = """ & ConvSPChars(EG1_export_group(E1_f_ln_info_biz_area_nm_to))   & """" & vbCr		
		Response.Write " End with				" & vbcr
		Response.Write "</Script>               " & vbCr
		
		Call DisplayMsgBox("990000", vbOKOnly, "", "", I_MKSCRIPT)
		
	End If

End Sub
%>
