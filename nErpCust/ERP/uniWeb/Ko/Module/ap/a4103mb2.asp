<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                 '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
'--------------- ������ coding part(��������,End)----------------------------------------------------------

    Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")         

																'��: ���� ��û�� ���� 

	Err.Clear																		'��: Protect system from crashing

	Dim PAPG015m					' ������ǥ���� ComProxy Dll ��� ���� 
	Dim strWkfg
	Dim ImportTransType

	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide


	' -- ���Ѱ����߰� 
	Const A386_I4_a_data_auth_data_BizAreaCd = 0
	Const A386_I4_a_data_auth_data_internal_cd = 1
	Const A386_I4_a_data_auth_data_sub_internal_cd = 2
	Const A386_I4_a_data_auth_data_auth_usr_id = 3

	Dim I4_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

	Redim I4_a_data_auth(3)
	I4_a_data_auth(A386_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I4_a_data_auth(A386_I4_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I4_a_data_auth(A386_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I4_a_data_auth(A386_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
				
	strWkfg = Request("htxtWorkFg")

	ImportTransType="AP003"

	Set PAPG015m = Server.CreateObject("PAPG015.cACnfmApSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Set PAPG015m = nothing
		Response.End	
	End If
		 
	Call PAPG015m.A_CONFIRM_AP_SVR (gStrGlobalCollection, ImportTransType, , Request("txtSpread"), iErrorPosition, I4_a_data_auth)
		    

	If CheckSYSTEMError2(Err, True, iErrorPosition & " ä����ȣ","","","","") = True Then
		Set PAPG015m = nothing		
		Response.End 
	End If
		    
	Set PAPG015m = nothing 

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbSaveOk          " & vbcr
	Response.Write "</Script>                 " & vbcr



%>