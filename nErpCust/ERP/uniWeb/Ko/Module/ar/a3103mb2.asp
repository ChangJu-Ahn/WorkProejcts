<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->

<%																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 


On Error Resume Next

Dim igTransType 
Dim iArrData
Dim iChangeOrgId
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

' -- ���Ѱ����߰� 
Const A050_I2_a_data_auth_data_BizAreaCd = 0
Const A050_I2_a_data_auth_data_internal_cd = 1
Const A050_I2_a_data_auth_data_sub_internal_cd = 2
Const A050_I2_a_data_auth_data_auth_usr_id = 3

Dim I2_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

Redim I2_a_data_auth(3)
I2_a_data_auth(A050_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
I2_a_data_auth(A050_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
I2_a_data_auth(A050_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
I2_a_data_auth(A050_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")    
		
	strMode = Request("txtMode")
	iChangeOrgId = Trim(request("OrgChangeId"))

	Err.Clear																		'��: Protect system from crashing

	If Request("txtMaxRows") = "" Then
	Call ServerMesgBox("700117",vbInformation, I_MKSCRIPT)              
	Response.End 
	End If

	LngMaxRow = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 

	igTransType = "AR005"
	iArrData = request("txtspread")

					
	Set pAr0015c = Server.CreateObject("PARG015.cACnfmArSvr")

	If CheckSYSTEMError(Err,True) = True Then
	Response.End 
	End If
	                         
	Call pAr0015c.A_CONFIRM_AR_SVR(gStrGlobalCollection, igTransType, iArrData, iErrorPosition, I2_a_data_auth)		
			
	If CheckSYSTEMError2(Err, True, iErrorPosition & " ä�ǹ�ȣ","","","","") = True Then
	Set pAr0015c = Nothing
	Response.End 
	End If    
	    
	Set pAr0015c = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbSaveOk          " & vbcr
	Response.Write "</Script>" & vbcr

%>
