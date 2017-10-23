
<%'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : Prepayment_Settlement
'*  3. Program ID        : f6102mb1
'*  4. Program 이름      : 선급금 청산 
'*  5. Program 설명      : 선급금 청산 List, Create, Delete, Update
'*  6. Complus 리스트    : 
'*  7. 최초 작성년월일   : 2000/10/07
'*  8. 최종 수정년월일   : 2002/11/15
'*  9. 최초 작성자       : 송봉훈 
'* 10. 최종 작성자       : Jeong Yong Kyun
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************

																	'☜ : ASP가 캐쉬되지 않도록 한다.
																	'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%																						'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

On Error Resume Next								 
Err.Clear 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim iPAFG610																			'☜:'입력/수정용 ComPlus Dll 사용 변수 
Dim iPrpaymPrpaymNo 
Dim iPrpaymSttlMentNo 
Dim iPrpaymSttlDt
Dim iPrpaymDocCur


Dim strMode																				'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngMaxRow3		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount 
Dim strData
Dim lgCurrency       


Dim iErrorPosition														'☜: Error 위치 
Dim iCommandSent		
Dim lgIntFlgMode
Dim lgSttlmentNo		

' -- 권한관리추가 
Const A725_I2_a_data_auth_data_BizAreaCd = 0
Const A725_I2_a_data_auth_data_internal_cd = 1
Const A725_I2_a_data_auth_data_sub_internal_cd = 2
Const A725_I2_a_data_auth_data_auth_usr_id = 3

		
	strMode = Request("txtMode")														'☜ : 현재 상태를 받음 
	Select Case strMode
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	' QUERY
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Case CStr(UID_M0001)																'☜: 현재 조회/Prev/Next 요청을 받음 
		Const R1_prpaym_no = 0
	    Const R1_prpaym_dt = 1
	    Const R1_ref_no = 2
	    Const R1_doc_cur = 3
	    Const R1_xch_rate = 4
	    Const R1_note_no = 5
	    Const R1_prpaym_amt = 6
	    Const R1_prpaym_loc_amt = 7
	    Const R1_sttl_amt = 8
	    Const R1_sttl_loc_amt = 9
	    Const R1_cls_amt = 10
	    Const R1_cls_loc_amt = 11
	    Const R1_bal_amt = 12
	    Const R1_bal_loc_amt = 13
	    Const R1_paym_type = 14
	    Const R1_prpaym_sts = 15
	    Const R1_conf_fg = 16
	    Const R1_gl_no = 17
	    Const R1_temp_gl_no = 18
	    Const R1_prpaym_desc = 19
	    Const R1_internal_cd = 20

	    Const R2_bp_cd = 0
	    Const R2_bp_nm = 1

	    Const R3_org_change_id = 0
	    Const R3_dept_cd = 1
	    Const R3_dept_nm = 2

	    Const RG1_sttl_no = 0
	    Const RG1_acct_cd = 1
	    Const RG1_acct_nm = 2
	    Const RG1_sttl_amt = 3
	    Const RG1_item_loc_amt = 4
	    Const RG1_sttl_loc_amt = 5
	    Const RG1_sttl_desc = 6

		Const R1_Sttl_prpaym_no = 0
		Const R1_Sttl_Doc_Cur = 1
		Const R1_Sttl_xch_rate = 2
		Const R1_Sttl_Dt = 3
	    Const R1_Sttl_gl_no = 4
		Const R1_Sttl_temp_gl_no = 5
		Const R1_Sttl_sttl_amt = 6
		Const R1_Sttl_item_loc_amt = 7
		

		Dim iarrRPrpaym
		Dim iarrRBizPartner
		Dim iarrRAcctDept
		Dim istrNextPrpaym
		Dim iarrRGSttl
		Dim iarrPrSttl		'//for leu project : 청산일자, 거래통화, 전표번호(결의, 회계)

		Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

		Redim I2_a_data_auth(3)

		I2_a_data_auth(A725_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
		I2_a_data_auth(A725_I2_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
		I2_a_data_auth(A725_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
		I2_a_data_auth(A725_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))


		lgStrPrevKey = Request("lgStrPrevKey")
		
	    Set iPAFG610 = Server.CreateObject("PAFG610.cFListPpSttlSvr")

	    '-----------------------------------------
	    'Com action result check area(OS,internal)
	    '-----------------------------------------
	    If CheckSYSTEMError(Err, True) = True Then					
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If   
	    '-----------------------------------------
	    'Data manipulate  area(import view match)
	    '-----------------------------------------
	    iPrpaymSttlMentNo = Trim(Request("txtSttlmentNo"))
	    
	    
	    '-----------------------------------------
	    'Com Action Area
	    '-----------------------------------------
		Call iPAFG610.F_LIST_PP_STTL_SVR(gStrGloBalCollection,iPrpaymSttlMentNo,iarrRPrpaym,iarrRBizPartner, _
												iarrRAcctDept,istrNextPrpaym,iarrPrSttl, iarrRGSttl, I2_a_data_auth)
	    '-----------------------------------------
	    'Com action result check area(OS,internal)
	    '-----------------------------------------
	    if CheckSYSTEMError(Err, True) = True Then					
			Set iPAFG605 = Nothing														'☜: ComPlus Unload
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If  
		'-----------------------------------------
		'Com action result check area(DB,internal)
		'-----------------------------------------

		LngMaxRow = Request("txtMaxRows")												'Save previous Maxrow                                                
		If isEmpty(iarrRGSttl) then 
			GroupCount = 0
		Else		
	   		GroupCount = UBound(iarrRGSttl,1) + 1
		End If   		
		
		If GroupCount > 0 Then
			If istrNextPrpaym = iarrRGSttl(GroupCount-1,RG1_sttl_no) Then
				StrNextKey = ""
			Else
				StrNextKey = istrNextPrpaym
			End If
		End If

		lgCurrency = ConvSPChars(iarrRPrpaym(R1_doc_cur))
		
		Response.Write " <Script Language=vbscript> " & vbCr
	    Response.Write " With parent                " & vbCr							'☜: 화면 처리 ASP 를 지칭함 
		Response.Write ".frm1.txtDeptCd.value		= """ & ConvSPChars(iarrRAcctDept(R3_dept_cd))         & """" & vbCr
		Response.Write ".frm1.txtDeptNm.value		= """ & ConvSPChars(iarrRAcctDept(R3_dept_nm))         & """" & vbCr
		Response.Write ".frm1.txtPrpaymDt.Text		= """ & UNIDateClientFormat(iarrRPrpaym(R1_prpaym_dt)) & """" & vbCr
		Response.Write ".frm1.txtBpCd.value			= """ & ConvSPChars(iarrRBizPartner(R2_bp_cd))         & """" & vbCr
		Response.Write ".frm1.txtBpNm.value			= """ & ConvSPChars(iarrRBizPartner(R2_bp_nm))         & """" & vbCr
		Response.Write ".frm1.txtRefNo.value		= """ & ConvSPChars(iarrRPrpaym(R1_ref_no))            & """" & vbCr
		Response.Write ".frm1.txtDocCur.value		= """ & ConvSPChars(iarrRPrpaym(R1_doc_cur))           & """" & vbCr
		Response.Write ".frm1.txtGlNo.value			= """ & ConvSPChars(iarrRPrpaym(R1_gl_no))             & """" & vbCr
		Response.Write ".frm1.txtTempGlNo.value		= """ & ConvSPChars(iarrRPrpaym(R1_temp_gl_no))             & """" & vbCr
		Response.Write ".frm1.txtXchRate.text		= """ & UNINumClientFormat(iarrRPrpaym(R1_xch_rate), ggExchRate.DecPoint, 0)                                             & """" & vbCr
		Response.Write ".frm1.txtPrpaymAmt.Text   	= """ & UNIConvNumDBToCompanyByCurrency(iarrRPrpaym(R1_prpaym_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                & """" & vbCr
		Response.Write ".frm1.txtPrpaymLocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(iarrRPrpaym(R1_prpaym_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
		Response.Write ".frm1.txtBalAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iarrRPrpaym(R1_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                   & """" & vbCr
		Response.Write ".frm1.txtBalLocAmt.Text   	= """ & UNIConvNumDBToCompanyByCurrency(iarrRPrpaym(R1_bal_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
		Response.Write ".frm1.txtPrpaymDesc.value	= """ & ConvSPChars(iarrRPrpaym(R1_prpaym_desc))       & """" & vbCr
		
		'//for leu project
		Response.Write ".frm1.hPrpaymNo.value		= """ & ConvSPChars(iarrPrSttl(R1_Sttl_prpaym_no)) & """" & vbCr
		Response.Write ".frm1.txtPrpaymNo.value		= """ & ConvSPChars(iarrPrSttl(R1_Sttl_prpaym_no)) & """" & vbCr
		Response.Write ".frm1.txtSttlDt.Text		= """ & UNIDateClientFormat(iarrPrSttl(R1_Sttl_Dt)) & """" & vbCr
		Response.Write ".frm1.txtSttlDocCur.value	= """ & ConvSPChars(iarrPrSttl(R1_Sttl_Doc_Cur))           & """" & vbCr
		Response.Write ".frm1.txtSttlTempGlNo.value	= """ & ConvSPChars(iarrPrSttl(R1_Sttl_temp_gl_no))             & """" & vbCr
		Response.Write ".frm1.txtSttlGlNo.value		= """ & ConvSPChars(iarrPrSttl(R1_Sttl_gl_no))             & """" & vbCr
		Response.Write ".frm1.txtSttlAmt.Text   	= """ & UNIConvNumDBToCompanyByCurrency(iarrPrSttl(R1_Sttl_sttl_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                & """" & vbCr
		Response.Write ".frm1.txtSttlLocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(iarrPrSttl(R1_Sttl_item_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
		'Response.Write ".frm1.txtSttlXchRate.text		= """ & UNINumClientFormat(iarrPrSttl(R1_Sttl_xch_rate), ggExchRate.DecPoint, 0)                                             & """" & vbCr
	


	'	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                

	  	For LngRow = 1 To GroupCount

	        strData = strData & Chr(11) & iarrRGSttl(LngRow-1,RG1_sttl_no)          	        '1  C_SttlNo
	        strData = strData & Chr(11) & ConvSPChars(iarrRGSttl(LngRow-1,RG1_acct_cd))			'3  C_AcctCd 
	        strData = strData & Chr(11) & ""													'4  C_AcctCdPopUp
	        strData = strData & Chr(11) & ConvSPChars(iarrRGSttl(LngRow-1,RG1_acct_nm))  		'5  C_AcctNm 
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iarrRGSttl(LngRow-1,RG1_sttl_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") 
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iarrRGSttl(LngRow-1,RG1_item_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iarrRGSttl(LngRow-1,RG1_sttl_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & ConvSPChars(iarrRGSttl(LngRow-1,RG1_sttl_desc))   
			strData = strData & Chr(11) & Cint(LngMaxRow) + Cint(LngRow)									
	        strData = strData & Chr(11) & Chr(12)
	    Next
		
			
		Response.Write ".ggoSpread.Source     = .frm1.vspdData                          " & vbCr
		Response.Write ".ggoSpread.SSShowData   """ & strData                      & """" & vbCr
		Response.Write ".lgStrPrevKey         = """ & ConvSPChars(StrNextKey)             & """" & vbCr
		Response.Write ".frm1.hSttlMentNo.value = """ & ConvSPChars(Request("txtSttlMentNo")) & """" & vbCr
		Response.Write " Call .DbQueryOk1                                                      " & vbCr
		Response.Write " End With                                                              " & vbCr
		Response.Write "</Script>	                                                           " & vbCr 
	    
	    Set iPAFG610 = Nothing
	  
	  
	  
	  
	    
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	' SAVE
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Case CStr(UID_M0002)																'☜: 저장 요청을 받음 

	    Err.Clear																		'☜: Protect system from crashing

		Dim I1_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  		Redim I1_a_data_auth(3)
		I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
		I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
		I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
		I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	    LngMaxRow  = CInt(Request("txtMaxRows"))										'☜: 최대 업데이트된 갯수 
	    LngMaxRow3 = CInt(Request("txtMaxRows3"))

	    Set iPAFG610 = Server.CreateObject("PAFG610.cFMngPpSttlSvr")
	    '-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If CheckSYSTEMError(Err, True) = True Then					
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If   
	    '-----------------------
	    'Data manipulate area
	    '-----------------------
		iPrpaymPrpaymNo = Trim(Request("hPrpaymNo"))
		
		'//FOR LEU PROJECT
		iPrpaymSttlMentNo = Trim(Request("hSttlmentNo")) 
		iPrpaymSttlDt = UNICONVDATE(Trim(Request("txtSttlDt")))
		iPrpaymDocCur = UCase(Trim(Request("txtSttlDocCur")))
		
		lgIntFlgMode = CInt(Request("txtFlgMode"))
		
		If lgIntFlgMode = OPMD_CMODE Then
			iCommandSent = "CREATE"
			lgSttlmentNo = iPAFG610.F_MANAGE_PP_STTL_SVR(gStrGloBalCollection,iCommandSent,iPrpaymPrpaymNo,iPrpaymSttlMentNo, _
		                                   iPrpaymSttlDt, iPrpaymDocCur,gCurrency, Request("txtSpread"),Request("txtSpread3"),iErrorPosition,I1_a_data_auth)
	    
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			iCommandSent = "UPDATE"
			lgSttlmentNo = iPAFG610.F_MANAGE_PP_STTL_SVR(gStrGloBalCollection,iCommandSent,iPrpaymPrpaymNo,iPrpaymSttlMentNo, _
		                                   iPrpaymSttlDt, iPrpaymDocCur,gCurrency, Request("txtSpread"),Request("txtSpread3"),iErrorPosition,I1_a_data_auth)
		End If
		
		
		'-----------------------------------------
	    'Com Action Area
	    '-----------------------------------------
		'-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If CheckSYSTEMError(Err, True) = True Then					
			Set iPAFG610 = Nothing														'☜: Unload Complus
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If   
	    
	'    if CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then					
	'		Set iPAFG605 = Nothing														'☜: ComPlus Unload
	'		Response.End																'☜: 비지니스 로직 처리를 종료함 
	'	End If  
		
		Set iPAFG610 = Nothing															'☜: Unload Complus

		Response.Write " <Script Language=VBScript> " & vbCr
		If Trim(ConvSPChars(lgSttlmentNo)) <> "" Then
			Response.Write "parent.frm1.txtSttlMentNo.value = """ & ConvSPChars(lgSttlmentNo) & """" & vbCr
			Response.Write "parent.frm1.hSttlMentNo.value = """ & ConvSPChars(lgSttlmentNo) & """" & vbCr
			Response.Write " Call parent.DbSaveOk       " & vbCr
		Else
			Response.Write " Call parent.DbDeleteOk       " & vbCr
		End If	
			
		Response.Write " </Script>                  " & vbCr
	
	
	
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	' DELETE
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Case CStr(UID_M0003)																'☜: 저장 요청을 받음 
	    Err.Clear																		'☜: Protect system from crashing

		Dim I3_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

		Redim I3_a_data_auth(3)

		I3_a_data_auth(A725_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
		I3_a_data_auth(A725_I2_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
		I3_a_data_auth(A725_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
		I3_a_data_auth(A725_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))


	    Set iPAFG610 = Server.CreateObject("PAFG610.cFMngPpSttlSvr")
	    '-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If CheckSYSTEMError(Err, True) = True Then					
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If   
	    '-----------------------
	    'Data manipulate area
	    '-----------------------
		iPrpaymPrpaymNo = Trim(Request("txtPrPaymNo"))
		
		'//FOR LEU PROJECT
		iPrpaymSttlMentNo = Trim(Request("txtSttlmentNo")) 
		
		lgIntFlgMode = CInt(Request("txtFlgMode"))
		
		iCommandSent = "DELETE"
		Call iPAFG610.F_MANAGE_PP_STTL_SVR(gStrGloBalCollection,iCommandSent,iPrpaymPrpaymNo,iPrpaymSttlMentNo, , , , , , , I3_a_data_auth)
	    
		
		
		'-----------------------------------------
	    'Com Action Area
	    '-----------------------------------------
		'-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If CheckSYSTEMError(Err, True) = True Then					
			Set iPAFG610 = Nothing														'☜: Unload Complus
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If   
	    
	'    if CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then					
	'		Set iPAFG605 = Nothing														'☜: ComPlus Unload
	'		Response.End																'☜: 비지니스 로직 처리를 종료함 
	'	End If  
		
		Set iPAFG610 = Nothing															'☜: Unload Complus

		Response.Write " <Script Language=VBScript> " & vbCr
		Response.Write " Call parent.DbDeleteOk       " & vbCr
		Response.Write " </Script>                  " & vbCr
	End Select
%>
