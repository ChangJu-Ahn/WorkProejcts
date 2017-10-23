<%
'**********************************************************************************************
'*  1. Module Name          : 입금반제(저장)
'*  2. Function Name        : 
'*  3. Program ID           : a3111mb2.adp
'*  4. Program Name         :	
'*  5. Program Desc         :
'*  6. Comproxy List        : +Ar0041r
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
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 

Call LoadBasisGlobalInf()

Dim pAr004m																	'조회용 ComProxy Dll 사용 변수 
Dim lgIntFlgMode
Dim LngMaxRow
Dim LngMaxRow0
Dim LngMaxRow1
Dim LngMaxRow3

' Com+ Conv. 변수 선언 
    
Dim iCommandSent 
Dim arrRowVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrVal																		'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim	lGrpCnt																		'☜: Group Count

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
Dim E1_b_monthly_exchange_rate
Dim E3_b_auto_numbering

'[CONVERSION INFORMATION]  View Name : import b_acct_dept
Const A366_I4_org_change_id = 0    
Const A366_I4_dept_cd = 1

'[CONVERSION INFORMATION]  View Name : import a_allc_rcpt
Const A366_I5_allc_no = 0    
Const A366_I5_allc_dt = 1
Const A366_I5_allc_type = 2
Const A366_I5_ref_no = 3
Const A366_I5_allc_amt = 4
Const A366_I5_allc_loc_amt = 5
Const A366_I5_dc_amt = 6
Const A366_I5_dc_loc_amt = 7
Const A366_I5_allc_rcpt_desc = 8
Const A366_I5_insrt_user_id = 9
Const A366_I5_updt_user_id = 10

	LngMaxRow = CInt(Request("txtMaxRows"))										'☜: 최대 업데이트된 갯수 
	LngMaxRow0 = CInt(Request("txtMaxRows0"))										'☜: 최대 업데이트된 갯수 
	LngMaxRow1 = CInt(Request("txtMaxRows1"))										'☜: 최대 업데이트된 갯수 
	LngMaxRow3 = CInt(Request("txtMaxRows3"))											'☜: 최대 업데이트된 갯수 
	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

	'-----------------------
	'Data manipulate area
	'-----------------------												    'Single 데이타 저장 
	ReDim I4_b_acct_dept(A366_I4_dept_cd)
	ReDim I5_a_allc_rcpt(A366_I5_updt_user_id)

	I1_a_acct_trans_type = "AR003"
	I6_b_currency = gCurrency
	I4_b_acct_dept(A366_I4_org_change_id) = UCase(Request("hOrgChangeId"))
	I4_b_acct_dept(A366_I4_dept_cd) = Trim(Request("txtDeptCd"))
	I3_a_allc_rcpt_assn = Trim(Request("txtDocCur"))
	I7_b_biz_partner = Trim(Request("txtBpCd"))
	I5_a_allc_rcpt(A366_I5_allc_no) = Trim(Request("txtAllcNo"))
	I5_a_allc_rcpt(A366_I5_allc_dt) = UNIConvDate(Request("txtAllcDt"))
	I5_a_allc_rcpt(A366_I5_allc_amt) = UNIConvNum(Request("txtAllcAmt"),0)
	I5_a_allc_rcpt(A366_I5_allc_loc_amt) = UNIConvNum(Request("txtAllcLocAmt"),0)
	I5_a_allc_rcpt(A366_I5_dc_amt) = UNIConvNum(Request("txtDcAmt"),0)
	I5_a_allc_rcpt(A366_I5_dc_loc_amt) = UNIConvNum(Request("txtDcLocAmt"),0)
	I5_a_allc_rcpt(A366_I5_allc_rcpt_desc) = Trim(Request("txtRcptDesc"))
	I5_a_allc_rcpt(A366_I5_allc_type) = "A"

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	lGrpCnt = 0

	If Request("txtSpread") <> "" Then
		importArray = Request("txtSpread")
	Else
		importArray = ""
		Call DisplayMsgBox("112500", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End	
	End If

	If Request("txtSpread1") <> "" Then
		importArray1 = Request("txtSpread1")
	Else
		importArray1 = ""
		Call DisplayMsgBox("112100", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End	
	End If

	If Request("txtSpread2") <> "" Then
		importArray2 = Request("txtSpread2")
	Else
		importArray2 = ""
	End If

	If Request("txtSpread3") <> "" Then
		importArray3 = Request("txtSpread3")
	Else
		importArray3 = ""
	End If

	If Trim(importArray) <> "" and Trim(importArray1) <> "" Then
		
		Set pAr004m = Server.CreateObject("PARG055.cAMntRcAllcSvr")
		
		If CheckSYSTEMError(Err,True) = True Then
			Response.End																			'☜: 비지니스 로직 처리를 종료함 
		End If		
				
		E3_b_auto_numbering = pAr004m.A_MAINT_RCPT_ALLC_SVR(gStrGlobalCollection,iCommandSent,I1_a_acct_trans_type,I2_a_acct,I3_a_allc_rcpt_assn, _ 
			importArray, I4_b_acct_dept,importArray1,importArray2,importArray3, _ 
			I5_a_allc_rcpt,I6_b_currency,I7_b_biz_partner)
		
		If CheckSYSTEMError(Err,True) = True Then
			Set pAr004m = Nothing																	'☜: ComProxy Unload
			Response.End																			'☜: 비지니스 로직 처리를 종료함 
		End If		
		
		Set pAr004m = Nothing																	'☜: ComProxy Unload		

	End IF

    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
    Response.Write " .DBSaveOk(""" & ConvSPChars(E3_b_auto_numbering)  & """)" & vbCr
    Response.Write "End With "					 & vbCr	  
    Response.Write "</Script>"           
%>
