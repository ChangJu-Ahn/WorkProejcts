<%
'**********************************************************************************************
'*  1. Module Name          : 선수금반제 
'*  2. Function Name        : 
'*  3. Program ID           : a3108mb3.aps
'*  4. Program Name         :	
'*  5. Program Desc         :
'*  6. Comproxy List        : +Ar0041pr
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2000/06/17
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : Chang Sung Hee
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
' 아래 함수는 비지니스 로직 시작되는 시점에서 호출해 주세요..
Call HideStatusWnd		
On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()

Dim pAr004m																	'☆ : 조회용 ComProxy Dll 사용 변수 

' Com+ Conv. 변수 선언 
    
Dim strGroup
Dim arrCount

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim iCommandSent
Dim I1_a_acct_trans_type
Dim I2_b_acct_dept
Dim I3_f_prrcpt
Dim I4_a_allc_rcpt
Dim I5_b_currency
Dim I6_b_biz_partner
Dim importArray
Dim importArray1
Dim importArray2
Dim E1_a_allc_rcpt

Const A365_I4_allc_no = 0    '[CONVERSION INFORMATION]  View Name : import a_allc_rcpt
Const A365_I4_allc_dt = 1
Const A365_I4_allc_type = 2
Const A365_I4_ref_no = 3
Const A365_I4_allc_amt = 4
Const A365_I4_allc_loc_amt = 5
Const A365_I4_dc_amt = 6
Const A365_I4_dc_loc_amt = 7
Const A365_I4_insrt_user_id = 8
Const A365_I4_updt_user_id = 9

	strMode = Request("txtMode")											'☜ : 현재 상태를 받음 

	If strMode = "" Then
		Response.End 
	ElseIf strMode <> CStr(UID_M0003) Then									'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
		Response.End 
	ElseIf Request("txtAllcNo") = "" Then									'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)			'조회 조건값이 비어있습니다!
		Response.End 
	End If

Dim I7_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 
Const A365_I7_a_data_auth_data_BizAreaCd = 0
Const A365_I7_a_data_auth_data_internal_cd = 1
Const A365_I7_a_data_auth_data_sub_internal_cd = 2
Const A365_I7_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I7_a_data_auth(3)
	I7_a_data_auth(A365_I7_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I7_a_data_auth(A365_I7_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I7_a_data_auth(A365_I7_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I7_a_data_auth(A365_I7_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))

	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	ReDim I4_a_allc_rcpt(A365_I4_updt_user_id)
	I4_a_allc_rcpt(A365_I4_allc_no) = Trim(Request("txtAllcNo"))
	I1_a_acct_trans_type	= "AR004"
	iCommandSent = "DELETE"

	importArray = ""
	importArray1 = ""
	importArray2 = ""
	'-----------------------
	'Com Action Area
	'-----------------------
	Set pAr004m = Server.CreateObject("PARG040.cAMntPrAllcSvr")
	
	If CheckSYSTEMError(Err,True) = True Then
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If	

	E1_a_allc_rcpt = pAr004m.A_MAINT_PRERCPT_ALLC_SVR(gStrGlobalCollection, iCommandSent, I1_a_acct_trans_type, , , I4_a_allc_rcpt, , , importArray, importArray1, importArray2,I7_a_data_auth)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set pAr004m = Nothing												'☜: ComProxy Unload
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	Set pAr004m = Nothing				

	Response.Write " <Script Language=vbscript> " & vbCr
	Response.write " Call parent.DbDeleteOk()   " & vbCr
	Response.Write " </Script>                  " & vbCr
%>
