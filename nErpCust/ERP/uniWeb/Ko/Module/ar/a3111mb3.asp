<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Open Ap
'*  3. Program ID           : a4101mb3
'*  4. Program Name         : Open Ap 삭제하는 Logic
'*  5. Program Desc         :
'*  6. Comproxy List        : +AP001M
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2000/04/10
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : You So Eun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************

								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()

Dim pAr004m																	'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim iCommandSent 

Dim I1_a_acct_trans_type
Dim I2_a_acct
Dim I3_a_allc_rcpt_assn
Dim importArray
Dim I4_b_acct_dept
Dim importArray1
Dim importArray2
Dim importArray3
Dim I5_a_allc_rcpt
Dim I6_b_currency
Dim I7_b_biz_partner

Const A366_I5_allc_no = 0    
Const A366_I5_allc_dt = 1
Const A366_I5_allc_type = 2
Const A366_I5_ref_no = 3
Const A366_I5_allc_amt = 4
Const A366_I5_allc_loc_amt = 5
Const A366_I5_dc_amt = 6
Const A366_I5_dc_loc_amt = 7
Const A366_I5_insrt_user_id = 8
Const A366_I5_updt_user_id = 9

	strMode = Request("txtMode")														'☜ : 현재 상태를 받음 

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then												'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
		Response.End 
		Call HideStatusWnd		
	ElseIf Request("txtAllcNo") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)						'조회 조건값이 비어있습니다!
		Response.End
		Call HideStatusWnd		 
	End If

Dim I8_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
Const A366_I8_a_data_auth_data_BizAreaCd = 0
Const A366_I8_a_data_auth_data_internal_cd = 1
Const A366_I8_a_data_auth_data_sub_internal_cd = 2
Const A366_I8_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I8_a_data_auth(3)
	I8_a_data_auth(A366_I8_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I8_a_data_auth(A366_I8_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I8_a_data_auth(A366_I8_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I8_a_data_auth(A366_I8_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	ReDim I5_a_allc_rcpt(A366_I5_updt_user_id)

	iCommandSent = "DELETE"
	I1_a_acct_trans_type = "AR003"
	I5_a_allc_rcpt(A366_I5_allc_no) = Trim(Request("txtAllcNo"))

	importArray  = ""
	importArray1 = ""
	importArray2 = ""
	importArray3 = ""

	Set pAr004m = Server.CreateObject("PARG055.cAMntRcAllcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																	'☜: 비지니스 로직 처리를 종료함 
	End If	
		
	Call pAr004m.A_MAINT_RCPT_ALLC_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,,,importArray, ,importArray1,importArray2,importArray3,I5_a_allc_rcpt,,I7_b_biz_partner,I8_a_data_auth)	
		
	If CheckSYSTEMError(Err,True) = True Then
		Set pAr004m = Nothing															'☜: ComProxy Unload
		Response.End																	'☜: 비지니스 로직 처리를 종료함 
	End If
	
	Set pAr004m = Nothing																'☜: Unload Comproxy	
	                                                
	Response.Write " <Script Language=vbscript> " & vbCr
   	Response.Write " Call parent.DbDeleteOk()   " & vbCr
	Response.Write " </Script>                  " & vbCr
%>
