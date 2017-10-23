<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a4103mb3
'*  4. Program Name         : 가수금내역삭제 
'*  5. Program Desc         : 가수금 내역을 삭제 
'*  6. Complus List         :
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : 조익성 
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************




'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	On Error Resume Next														'☜: 

	Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	
	Call LoadBasisGlobalInf()

	Dim iPARG020																'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
		
	Dim iArrSpread
	Dim ImportTransType
	Dim ImportDocCur
	Dim ImportPartnerBpCd
	Dim ImportDeptCd
	Dim ImportARcpt

	Dim ExportAGl
	Dim ExportRcpt

	Const RcptNo				= 0
	Const RcptDt				= 1
	Const DocCur				= 2
	Const XchRate				= 3
	Const BnkChgAmt				= 4
	Const BnkChgLocAmt			= 5
	Const RcptType				= 6
	Const RcptDesc				= 7
	Const RefNo					= 8

	'//기초치를 위한FLAG
	Const Gl_Input_Type			= 9		'//기초치구분 
	Const GlFlag				= 10	'//전표구분 

	ReDim ImportARcpt(GlFlag)

	Redim ImportDeptCd(2)

	' -- 권한관리추가 
	Const A114_I11_a_data_auth_data_BizAreaCd = 0
	Const A114_I11_a_data_auth_data_internal_cd = 1
	Const A114_I11_a_data_auth_data_sub_internal_cd = 2
	Const A114_I11_a_data_auth_data_auth_usr_id = 3

	Dim I11_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

	Redim I11_a_data_auth(3)
	I11_a_data_auth(A114_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

	If strMode = "" Then
		Response.End 
	ElseIf strMode <> CStr(UID_M0003) Then										'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
		Response.End 
	ElseIf Request("txtRcptNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If

	ImportTransType				= "AR001"
	ImportARcpt(RcptNo)			= Trim(Request("txtRcptNo"))

	ImportARcpt(Gl_Input_Type)	= "RP"		'//가수금등록 
	


	Set iPARG020 = Server.CreateObject("PARG020.cAMngRcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	                         
	Call iPARG020.MANAGE_RCPT_SVR(gStrGlobalCollection, "DELETE", ImportTransType, gCurrency, _
		                                       ImportARcpt, ImportPartnerBpCd, ImportDeptCd, iArrSpread, _
		                                       ExportAGl, ExportRcpt, I11_a_data_auth)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG020 = Nothing		
		Response.End 
	End If
	    
	Set iPARG020 = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.DbDeleteOk()      " & vbcr
	Response.Write "</Script>" & vbcr
%>
