<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->

<%																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 


On Error Resume Next

Dim igTransType 
Dim iArrData
Dim iChangeOrgId
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

' -- 권한관리추가 
Const A050_I2_a_data_auth_data_BizAreaCd = 0
Const A050_I2_a_data_auth_data_internal_cd = 1
Const A050_I2_a_data_auth_data_sub_internal_cd = 2
Const A050_I2_a_data_auth_data_auth_usr_id = 3

Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

Redim I2_a_data_auth(3)
I2_a_data_auth(A050_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
I2_a_data_auth(A050_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
I2_a_data_auth(A050_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
I2_a_data_auth(A050_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

	Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")    
		
	strMode = Request("txtMode")
	iChangeOrgId = Trim(request("OrgChangeId"))

	Err.Clear																		'☜: Protect system from crashing

	If Request("txtMaxRows") = "" Then
	Call ServerMesgBox("700117",vbInformation, I_MKSCRIPT)              
	Response.End 
	End If

	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

	igTransType = "AR005"
	iArrData = request("txtspread")

					
	Set pAr0015c = Server.CreateObject("PARG015.cACnfmArSvr")

	If CheckSYSTEMError(Err,True) = True Then
	Response.End 
	End If
	                         
	Call pAr0015c.A_CONFIRM_AR_SVR(gStrGlobalCollection, igTransType, iArrData, iErrorPosition, I2_a_data_auth)		
			
	If CheckSYSTEMError2(Err, True, iErrorPosition & " 채권번호","","","","") = True Then
	Set pAr0015c = Nothing
	Response.End 
	End If    
	    
	Set pAr0015c = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbSaveOk          " & vbcr
	Response.Write "</Script>" & vbcr

%>
