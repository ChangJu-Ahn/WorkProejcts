<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf()
Call HideStatusWnd

On Error Resume Next														'��: 

Dim iPAVG015	                 				                                '�� : �Է�/������ ComProxy Dll ��� ���� 
Dim lgOpModeCRUD 																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim I1_b_biz_partner
Dim I2_vat_from_Dt
Dim I3_vat_to_Dt 
Dim I4_vat_Updt_UserId


lgOpModeCRUD      = Request("txtMode")											'��: Read Operation Mode (CRUD)
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)														
        'Call SubBizQuery()														
        'Call SubBizQueryMulti()												
    Case CStr(UID_M0002)														
        'Call SubBizSave()														
         Call SubUpdRgstNo()														
    Case CStr(UID_M0003)														
        'Call SubBizDelete()													
End Select

Response.End 


Sub SubUpdRgstNo()

    On Error Resume Next														'��: 
    Err.Clear                                                               '��: Protect system from crashing        
	'********************************************************  
	'                        Execution
	'********************************************************  
	
    Set iPAVG015 = Server.CreateObject("PAVG015.cAExecVatRegNoSvr")       
    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
   If CheckSYSTEMError(Err, True) = True Then					
         Set iPAVG015 = Nothing
       Exit Sub
       
    End If    

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_b_biz_partner  = UCase(Trim(Request("txtBpCd")))
    I2_vat_from_Dt = UNIConvDate(Request("txtFromDt"))
    I3_vat_to_Dt   = UNIConvDate(Request("txtToDt"))
    I4_vat_Updt_UserId   = gUsrID
    
        
    Call iPAVG015.AG010M_EXECUTE_VAT_RGST_NO_SVR(gStrGlobalCollection,I1_b_biz_partner,I2_vat_from_Dt,I3_vat_to_Dt,I4_vat_Updt_UserId )
	
	'Response.Write CheckSYSTEMError(Err, True)
	If CheckSYSTEMError(Err, True) = True Then					
         Set iPAVG015 = Nothing
       Exit Sub
    else 
		
		 Call DisplayMsgBox("115123", vbInformation, "", "", I_MKSCRIPT)  ' msgno �߰��Ǹ���~ 
    End If    


Set iPAVG015 = Nothing
End Sub
%>
