<%
'**********************************************************************************************
'*  1. Module Name          : open ap 삭제 
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +AP001M
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 1999/09/10
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Mrs Kim
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
<%													

Call HideStatusWnd		
On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()

Dim pAr004m																	'☆ : 조회용 ComProxy Dll 사용 변수 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim strGroup
Dim arrCount

Dim iCommandSent 

Dim I1_a_acct_trans_type
Dim I2_a_acct
Dim I3_a_allc_rcpt_assn
Dim IG1_import_group_rcpt
Dim I4_b_acct_dept
Dim I5_a_allc_rcpt
Dim I6_b_currency
Dim I7_b_biz_partner
Dim importArray

Const A358_I5_allc_no = 0    '[CONVERSION INFORMATION]  View Name : import a_allc_rcpt
Const A358_I5_allc_dt = 1
Const A358_I5_allc_type = 2
Const A358_I5_ref_no = 3
Const A358_I5_allc_amt = 4
Const A358_I5_allc_loc_amt = 5
Const A358_I5_dc_amt = 6
Const A358_I5_dc_loc_amt = 7
Const A358_I5_insrt_user_id = 8
Const A358_I5_updt_user_id = 9

	ReDim IG1_import_group_rcpt(0,A358_IG1_I3_updt_dt)
	ReDim I4_b_acct_dept(A358_I4_dept_cd)
	ReDim I5_a_allc_rcpt(A358_I5_updt_user_id)

Dim I8_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
Const A358_I8_a_data_auth_data_BizAreaCd = 0
Const A358_I8_a_data_auth_data_internal_cd = 1
Const A358_I8_a_data_auth_data_sub_internal_cd = 2
Const A358_I8_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I8_a_data_auth(3)
	I8_a_data_auth(A358_I8_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I8_a_data_auth(A358_I8_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I8_a_data_auth(A358_I8_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I8_a_data_auth(A358_I8_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	strMode = Request("txtlgMode")												'☜ : 현재 상태를 받음 

	If strMode = "" Then	
		Response.End 
	ElseIf strMode <> CStr(UID_M0003) Then										'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함	
		Response.End 
	ElseIf Request("txtAllcNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)				'조회 조건값이 비어있습니다!
		Response.End 
	End If

	I5_a_allc_rcpt(A358_I5_allc_no) = Trim(Request("txtAllcNo"))
	importArray = ""
	iCommandSent = "DELETE"
	I1_a_acct_trans_type = "AR006"
	I2_a_acct = ""
	I3_a_allc_rcpt_assn = ""
	I6_b_currency = ""
	I7_b_biz_partner = ""

	Set pAr004m = Server.CreateObject("PARG080.cAMntAllcRcByApSvr")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pAr004m = Nothing													'☜: ComProxy Unload
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:
		Response.End															'☜: 비지니스 로직 처리를 종료함 
	End If

	E2_b_auto_numbering = pAr004m.A_MAINT_ALLC_RCPT_BY_AP_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,,,,, I5_a_allc_rcpt,,, importArray,I8_a_data_auth)

	If CheckSYSTEMError(Err,True) = True Then
		Set pAr004m = Nothing													'☜: ComProxy Unload
		Response.End															'☜: 비지니스 로직 처리를 종료함 
	End If

	Set pAr004m = Nothing

	Response.Write " <Script Language=vbscript>  " & vbCr
	Response.Write "	Call parent.DbDeleteOk() " & vbCr
	Response.Write " </Script>                   " & vbCr
%>	
