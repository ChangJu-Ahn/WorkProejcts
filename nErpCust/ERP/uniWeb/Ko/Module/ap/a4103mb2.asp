<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                 '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")         

																'☜: 저장 요청을 받음 

	Err.Clear																		'☜: Protect system from crashing

	Dim PAPG015m					' 결의전표승인 ComProxy Dll 사용 변수 
	Dim strWkfg
	Dim ImportTransType

	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide


	' -- 권한관리추가 
	Const A386_I4_a_data_auth_data_BizAreaCd = 0
	Const A386_I4_a_data_auth_data_internal_cd = 1
	Const A386_I4_a_data_auth_data_sub_internal_cd = 2
	Const A386_I4_a_data_auth_data_auth_usr_id = 3

	Dim I4_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

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
		    

	If CheckSYSTEMError2(Err, True, iErrorPosition & " 채무번호","","","","") = True Then
		Set PAPG015m = nothing		
		Response.End 
	End If
		    
	Set PAPG015m = nothing 

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.dbSaveOk          " & vbcr
	Response.Write "</Script>                 " & vbcr



%>