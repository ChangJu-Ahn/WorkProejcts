<%
'**********************************************************************************************
'*  1. Module Name          : 선수금반제 
'*  2. Function Name        : 
'*  3. Program ID           : a3108mb2.aps
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




'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
' 아래 함수는 비지니스 로직 시작되는 시점에서 호출해 주세요..
Call HideStatusWnd		
On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()													

Dim pAr004m																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim IntRows
Dim IntCols
Dim vbIntRet
Dim lEndRow
Dim boolCheck
Dim lgIntFlgMode
Dim LngMaxRow
Dim LngMaxRow1
Dim LngMaxRow3

' Com+ Conv. 변수 선언 
    
Dim strGroup
Dim arrCount

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

' 첨자선언 
Const A365_I2_org_change_id = 0    '[CONVERSION INFORMATION]  View Name : import b_acct_dept
Const A365_I2_dept_cd = 1

Const A365_I3_prrcpt_no = 0    '[CONVERSION INFORMATION]  View Name : import f_prrcpt
Const A365_I3_prrcpt_dt = 1
Const A365_I3_doc_cur = 2

Const A365_I4_allc_no = 0    '[CONVERSION INFORMATION]  View Name : import a_allc_rcpt
Const A365_I4_allc_dt = 1
Const A365_I4_allc_type = 2
Const A365_I4_ref_no = 3
Const A365_I4_allc_amt = 4
Const A365_I4_allc_loc_amt = 5
Const A365_I4_dc_amt = 6
Const A365_I4_dc_loc_amt = 7
Const A365_I4_allc_rcpt_desc = 8
Const A365_I4_insrt_user_id = 9
Const A365_I4_updt_user_id = 10

	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
	LngMaxRow1 = CInt(Request("txtMaxRows1"))										'☜: 최대 업데이트된 갯수 
	LngMaxRow3 = CInt(Request("txtMaxRows3"))										'☜: 최대 업데이트된 갯수 
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

	ReDim I2_b_acct_dept(A365_I2_dept_cd)
	ReDim I3_f_prrcpt(A365_I3_doc_cur)
	ReDim I4_a_allc_rcpt(A365_I4_updt_user_id)

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
	'Data manipulate area
	'-----------------------				
	I1_a_acct_trans_type					= "AR004"

	I2_b_acct_dept(A365_I2_org_change_id)	= GetGlobalInf("gChangeOrgId")
	I2_b_acct_dept(A365_I2_dept_cd)			= Trim(Request("txtDeptCd"))

	I3_f_prrcpt(A365_I3_prrcpt_no)			= Trim(Request("txtPrNo"))
	I3_f_prrcpt(A365_I3_prrcpt_dt)			= UNIConvDate(Request("txtPrDt"))
	I3_f_prrcpt(A365_I3_doc_cur)			= UCase(Trim(Request("txtDocCur")))
	
	I4_a_allc_rcpt(A365_I4_allc_no)			= Trim(Request("htxtAllcNo"))
	I4_a_allc_rcpt(A365_I4_allc_dt)			= UNIConvDate(Request("txtAllcDt"))
	I4_a_allc_rcpt(A365_I4_allc_type)		= "P"
	I4_a_allc_rcpt(A365_I4_allc_amt)		= UNIConvNum(Request("txtClsAmt"),0)
	I4_a_allc_rcpt(A365_I4_allc_loc_amt)	= UNIConvNum(Request("txtClsLocAmt"),0)
	I4_a_allc_rcpt(A365_I4_dc_amt)			= 0
	I4_a_allc_rcpt(A365_I4_dc_loc_amt)		= 0
	I4_a_allc_rcpt(A365_I4_allc_rcpt_desc)	= Request("txtDesc")
	I4_a_allc_rcpt(A365_I4_insrt_user_id)	= Request("txtUpdtUserId")
	I4_a_allc_rcpt(A365_I4_updt_user_id)	= Request("txtUpdtUserId")
	
	I5_b_currency							= gCurrency
	I6_b_biz_partner						= Trim(Request("txtBizCd"))

	If lgIntFlgMode = OPMD_CMODE Then	
		iCommandSent = "CREATE"	
	ElseIf lgIntFlgMode = OPMD_UMODE Then	
		iCommandSent = "UPDATE"
	End If

	If Trim(Request("txtPrNo")) = "" Then
		Call DisplayMsgBox("112124", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If

	If Request("txtSpread") = "" Then
		Call DisplayMsgBox("112100", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	Else
		importArray = Request("txtSpread")
	
		If Request("txtSpread1") <> "" Then
			importArray1 = Request("txtSpread1")
			
			If Request("txtSpread3") <> "" Then
				importArray2 = Request("txtSpread3")
			Else
				importArray2 = ""
			End If
		Else
			importArray1 = ""
			
			If Request("txtSpread3") <> "" Then
				importArray2 = Request("txtSpread3")
			Else
				importArray2 = ""
			End If
		End If	

		Set pAr004m = Server.CreateObject("PARG040.cAMntPrAllcSvr")
	
		If CheckSYSTEMError(Err,True) = True Then
			Response.End																			'☜: 비지니스 로직 처리를 종료함 
		End If			

		E1_a_allc_rcpt = pAr004m.A_MAINT_PRERCPT_ALLC_SVR(gStrGlobalCollection, iCommandSent, I1_a_acct_trans_type, I2_b_acct_dept, _
			I3_f_prrcpt, I4_a_allc_rcpt, I5_b_currency, I6_b_biz_partner, importArray, importArray1, importArray2,I7_a_data_auth)

		If CheckSYSTEMError(Err,True) = True Then
			Set pAr004m = Nothing																	'☜: ComProxy Unload
			Response.End																			'☜: 비지니스 로직 처리를 종료함 
		End If		
	
		Set pAr004m = Nothing
		
	End If

	Response.write " <Script Language=vbscript> " & vbCR
	Response.write " With parent		        " & vbCr
	If  E1_a_allc_rcpt <> "" Then																	'☜: 화면 처리 ASP 를 지칭함 
		Response.write " .frm1.txtAllcNo.value = """ & E1_a_allc_rcpt & """" & vbCr
	End if		
	Response.write " .DbSaveOk """ & E1_a_allc_rcpt & """" & vbCr
	Response.write " End With  " & vbCr
	Response.Write " </Script> " & vbCr
%>
