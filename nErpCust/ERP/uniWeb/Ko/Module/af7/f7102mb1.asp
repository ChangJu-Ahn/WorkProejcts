
<%'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : PreReceipt
'*  3. Program ID        : f7102mb1
'*  4. Program 이름      : 선수금 청산 
'*  5. Program 설명      : 선수금 청산 리스트 조회 , 청산추가 , 삭제 , 수정 
'*  6. Comproxy 리스트   : fr0021 , fr0028
'*  7. 최초 작성년월일   : 2000/10/7
'*  8. 최종 수정년월일   : 2002/11/19
'*  9. 최초 작성자       : 송봉훈 
'* 10. 최종 작성자       : Jeong Yong Kyun
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'**********************************************************************************************
														'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True														'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

On Error Resume Next														'☜: 
Err.Clear 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim StrNextKey																' 다음 값 
Dim lgStrPrevKey															' 이전 값 
Dim LngMaxRow																' 현재 그리드의 최대Row
Dim LngMaxRow3																' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount 
Dim LngLastRow      
Dim strData
Dim lgCurrency       

Dim iPAFG710																'입력/수정용 ComPlus Dll 사용 변수 
Dim iprrcptprrcptNo
Dim iprrcptSttlMentNo 
Dim iprrcptSttlDt
Dim iprrcptDocCur

Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim	lGrpCnt																		'☜: Group Count
Dim strCode																		'Lookup 용 리턴 변수 
Dim iErrorPosition	
Dim iCommandSent		
Dim lgIntFlgMode
Dim lgSttlmentNo
			
	

Dim igCurrency

strmode= request("txtmode")

Select Case strMode
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' QUERY
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

	Const R1_prrcpt_no = 0
    Const R1_prrcpt_dt = 1
    Const R1_ref_no = 2
    Const R1_doc_cur = 3
    Const R1_xch_rate = 4
    Const R1_prrcpt_amt = 5
    Const R1_prrcpt_loc_amt = 6
    Const R1_sttl_amt = 7
    Const R1_sttl_loc_amt = 8
    Const R1_cls_amt = 9
    Const R1_cls_loc_amt = 10
    Const R1_bal_amt = 11
    Const R1_bal_loc_amt = 12
    Const R1_paym_type = 13
    Const R1_prrcpt_sts = 14
    Const R1_conf_fg = 15
    Const R1_gl_no = 16
    Const R1_temp_gl_no = 17
    Const R1_prrcpt_desc = 18
    Const R1_internal_cd = 19

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

	Const R1_sttl_prrcpt_no = 0
	Const R1_Sttl_Doc_Cur = 1
	Const R1_Sttl_xch_rate = 2
	Const R1_Sttl_Dt = 3
	Const R1_Sttl_gl_no = 4
	Const R1_Sttl_temp_gl_no = 5
	Const R1_Sttl_sttl_amt = 6
	Const R1_Sttl_item_loc_amt = 7
	
	Dim iarrRprrcpt
	Dim iarrRBizPartner
	Dim iarrRAcctDept
	Dim istrNextprrcpt
	Dim iarrRGSttl
	Dim iarrRpSttl		'//for leu project : 청산일자, 거래통화, 전표번호(결의, 회계)

	' -- 조회용 
	' -- 권한관리추가 
	Const A838_I2_a_data_auth_data_BizAreaCd = 0
	Const A838_I2_a_data_auth_data_internal_cd = 1
	Const A838_I2_a_data_auth_data_sub_internal_cd = 2
	Const A838_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A838_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I2_a_data_auth(A838_I2_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I2_a_data_auth(A838_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I2_a_data_auth(A838_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

	
	lgStrPrevKey = Request("lgStrPrevKey")
	
    Set iPAFG710 = Server.CreateObject("PAFG710.cFListPrSttlSvr")

    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End								'☜: 비지니스 로직 처리를 종료함 
	End If   

    '-----------------------------------------
    'Data manipulate  area(import view match)
    '-----------------------------------------
    iprrcptSttlMentNo = Trim(Request("txtSttlmentNo"))

    '-----------------------------------------
    'Com Action Area
    '-----------------------------------------
	Call iPAFG710.F_LIST_PR_STTL_SVR(gStrGloBalCollection,iprrcptSttlMentNo,iarrRprrcpt,iarrRBizPartner, _
											iarrRAcctDept,istrNextprrcpt,iarrRpSttl,iarrRGSttl,I2_a_data_auth)
 
	
    '-----------------------------------------
    'Com action result check area(OS,internal)
    '-----------------------------------------
    if CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG710 = Nothing						'☜: ComPlus Unload
		Response.End								'☜: 비지니스 로직 처리를 종료함 
	End If  

	'-----------------------------------------
	'Com action result check area(DB,internal)
	'-----------------------------------------

	LngMaxRow = Request("txtMaxRows")										'Save previous Maxrow                                                
	if isEmpty(iarrRGSttl) then 
		GroupCount = 0
	else		
   		GroupCount = UBound(iarrRGSttl,1) + 1
	end if   		
	
	If GroupCount > 0 Then
		If istrNextprrcpt = iarrRGSttl(GroupCount-1,RG1_sttl_no) Then
			StrNextKey = ""
		Else
			StrNextKey = istrNextprrcpt
		End If
	End If

	lgCurrency = ConvSPChars(iarrRprrcpt(R1_doc_cur))
	
	Response.Write " <Script Language=vbscript> " & vbCr
    Response.Write " With parent                " & vbCr										'☜: 화면 처리 ASP 를 지칭함 
	
	Response.Write ".frm1.txtDeptCd.value		= """ & ConvSPChars(iarrRAcctDept(R3_dept_cd))         & """" & vbCr
	Response.Write ".frm1.txtDeptNm.value		= """ & ConvSPChars(iarrRAcctDept(R3_dept_nm))         & """" & vbCr
	Response.Write ".frm1.txtprrcptDt.Text		= """ & UNIDateClientFormat(iarrRprrcpt(R1_prrcpt_dt)) & """" & vbCr
	Response.Write ".frm1.txtBpCd.value			= """ & ConvSPChars(iarrRBizPartner(R2_bp_cd))         & """" & vbCr
	Response.Write ".frm1.txtBpNm.value			= """ & ConvSPChars(iarrRBizPartner(R2_bp_nm))         & """" & vbCr
	Response.Write ".frm1.txtRefNo.value		= """ & ConvSPChars(iarrRprrcpt(R1_ref_no))            & """" & vbCr
	Response.Write ".frm1.txtDocCur.value		= """ & ConvSPChars(iarrRprrcpt(R1_doc_cur))           & """" & vbCr
	Response.Write ".frm1.txtXchRate.text		= """ & UNINumClientFormat(iarrRprrcpt(R1_xch_rate), ggExchRate.DecPoint, 0)                                             & """" & vbCr
	Response.Write ".frm1.txtGlNo.value			= """ & ConvSPChars(iarrRprrcpt(R1_gl_no))            & """" & vbCr
	Response.Write ".frm1.txtTempGlNo.value		= """ & ConvSPChars(iarrRprrcpt(R1_temp_gl_no))           & """" & vbCr
	Response.Write ".frm1.txtprrcptAmt.value	= """ & UNIConvNumDBToCompanyByCurrency(iarrRprrcpt(R1_prrcpt_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                & """" & vbCr
	Response.Write ".frm1.txtprrcptLocAmt.value	= """ & UNIConvNumDBToCompanyByCurrency(iarrRprrcpt(R1_prrcpt_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
	Response.Write ".frm1.txtBalAmt.value		= """ & UNIConvNumDBToCompanyByCurrency(iarrRprrcpt(R1_bal_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                   & """" & vbCr
	Response.Write ".frm1.txtBalLocAmt.value	= """ & UNIConvNumDBToCompanyByCurrency(iarrRprrcpt(R1_bal_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")     & """" & vbCr
		
	Response.Write ".frm1.txtprrcptDesc.value	= """ & ConvSPChars(iarrRprrcpt(R1_prrcpt_desc))       & """" & vbCr
	
'	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                

  	'//for leu project
	Response.Write ".frm1.htxtPrrcptNo.value		= """ & ConvSPChars(iarrRpSttl(R1_sttl_prrcpt_no))         & """" & vbCr
	Response.Write ".frm1.txtPrrcptNo.value		= """ & ConvSPChars(iarrRpSttl(R1_sttl_prrcpt_no))         & """" & vbCr
	Response.Write ".frm1.txtSttlDt.Text		= """ & UNIDateClientFormat(iarrRpSttl(R1_Sttl_Dt)) & """" & vbCr
	Response.Write ".frm1.txtSttlDocCur.value	= """ & ConvSPChars(iarrRpSttl(R1_Sttl_Doc_Cur))           & """" & vbCr
	Response.Write ".frm1.txtSttlTempGlNo.value	= """ & ConvSPChars(iarrRpSttl(R1_Sttl_temp_gl_no))             & """" & vbCr
	Response.Write ".frm1.txtSttlGlNo.value		= """ & ConvSPChars(iarrRpSttl(R1_Sttl_gl_no))             & """" & vbCr
	Response.Write ".frm1.txtSttlAmt.Text   	= """ & UNIConvNumDBToCompanyByCurrency(iarrRpSttl(R1_Sttl_sttl_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")                & """" & vbCr
	Response.Write ".frm1.txtSttlLocAmt.Text	= """ & UNIConvNumDBToCompanyByCurrency(iarrRpSttl(R1_Sttl_item_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr
  	'Response.Write ".frm1.txtSttlXchRate.text	= """ & UNINumClientFormat(iarrRpSttl(R1_Sttl_xch_rate), ggExchRate.DecPoint, 0)                                             & """" & vbCr
  	
  	
  	For LngRow = 1 To GroupCount

        strData = strData & Chr(11) & ConvSPChars(iarrRGSttl(LngRow-1,RG1_sttl_no))	        '1  C_SttlNo
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

	Response.Write ".ggoSpread.Source        = .frm1.vspdData                          " & vbCr
	Response.Write ".ggoSpread.SSShowData      """ & strData                      & """" & vbCr
	Response.Write ".lgStrPrevKey            = """ & ConvSPChars(StrNextKey)             & """" & vbCr
	Response.Write ".frm1.hSttlmentNo.value = """ & ConvSPChars(Request("txtSttlmentNo")) & """" & vbCr
	Response.Write " Call .DbQueryOk1                                                         " & vbCr
	Response.Write " End With                                                                 " & vbCr
	Response.Write "</Script>	                                                              " & vbCr
    
    Set iPAFG710 = Nothing

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' SAVE
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 

	' -- 저장용 
	' -- 권한관리추가 
	Const A837_I6_a_data_auth_data_BizAreaCd = 0
	Const A837_I6_a_data_auth_data_internal_cd = 1
	Const A837_I6_a_data_auth_data_sub_internal_cd = 2
	Const A837_I6_a_data_auth_data_auth_usr_id = 3

	Dim I6_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I6_a_data_auth(3)
	I6_a_data_auth(A837_I6_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I6_a_data_auth(A837_I6_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I6_a_data_auth(A837_I6_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I6_a_data_auth(A837_I6_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
												
    Err.Clear																		'☜: Protect system from crashing

    LngMaxRow  = CInt(Request("txtMaxRows"))										'☜: 최대 업데이트된 갯수 
    LngMaxRow3 = CInt(Request("txtMaxRows3"))

    Set iPAFG710 = Server.CreateObject("PAFG710.cFMngPrSttlSvr")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End																'☜: 비지니스 로직 처리를 종료함 
	End If   

    '-----------------------
    'Data manipulate area
    '-----------------------
    
	iprrcptprrcptNo = Trim(Request("htxtPrrcptNo"))

	'//FOR LEU PROJECT
	iprrcptSttlMentNo = Trim(Request("hSttlmentNo")) 
	iprrcptSttlDt = uniconvdate(Trim(Request("txtSttlDt")))
	iprrcptDocCur = UCase(Trim(Request("txtSttlDocCur")))
	lgIntFlgMode = CInt(Request("txtFlgMode"))
		
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
		lgSttlmentNo =  iPAFG710.F_MANAGE_PR_STTL_SVR(gStrGloBalCollection,iCommandSent,iprrcptprrcptNo,iprrcptSttlMentNo, _
	                                   iprrcptSttlDt, iprrcptDocCur,gCurrency, Request("txtSpread"),Request("txtSpread3"),iErrorPosition,I6_a_data_auth)

	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
		lgSttlmentNo =  iPAFG710.F_MANAGE_PR_STTL_SVR(gStrGloBalCollection,iCommandSent,iprrcptprrcptNo,iprrcptSttlMentNo, _
	                                   iprrcptSttlDt, iprrcptDocCur,gCurrency, Request("txtSpread"),Request("txtSpread3"),iErrorPosition,I6_a_data_auth)
	End If
		
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then					
    	Set iPAFG710 = Nothing														'☜: ComPlus Unload
		Response.End																'☜: 비지니스 로직 처리를 종료함 
	End If   
    
'   if CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then					
'		Set iPAFG710 = Nothing														'☜: ComPlus Unload
'		Response.End																'☜: 비지니스 로직 처리를 종료함 
'	End If  
	
	Set iPAFG710 = Nothing															'☜: Unload Complus

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
' delete
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Case CStr(UID_M0003)																'☜: 저장 요청을 받음 

		Err.Clear																		'☜: Protect system from crashing

	' -- 조회용 
	' -- 권한관리추가 
	Const A697_I6_a_data_auth_data_BizAreaCd = 0
	Const A697_I6_a_data_auth_data_internal_cd = 1
	Const A697_I6_a_data_auth_data_sub_internal_cd = 2
	Const A697_I6_a_data_auth_data_auth_usr_id = 3

	'Dim I6_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I6_a_data_auth(3)
	I6_a_data_auth(A697_I6_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I6_a_data_auth(A697_I6_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I6_a_data_auth(A697_I6_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I6_a_data_auth(A697_I6_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
	    Set iPAFG710 = Server.CreateObject("PAFG710.cFMngPrSttlSvr")

	    '-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    
	    If CheckSYSTEMError(Err, True) = True Then					
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If   

	    '-----------------------
	    'Data manipulate area
	    '-----------------------
	    
		iprrcptprrcptNo = Trim(Request("txtPrRcptNo"))

		'//FOR LEU PROJECT
		iprrcptSttlMentNo = Trim(Request("txtSttlmentNo")) 
		lgIntFlgMode = CInt(Request("txtFlgMode"))
			
		iCommandSent = "DELETE"
		Call iPAFG710.F_MANAGE_PR_STTL_SVR(gStrGloBalCollection,iCommandSent,iprrcptprrcptNo,iprrcptSttlMentNo, , , , , , ,I6_a_data_auth)
	    
	    '-----------------------
	    'Com action result check area(OS,internal)
	    '-----------------------
	    If CheckSYSTEMError(Err, True) = True Then					
	    	Set iPAFG710 = Nothing														'☜: ComPlus Unload
			Response.End																'☜: 비지니스 로직 처리를 종료함 
		End If   
	    
	'   if CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then					
	'		Set iPAFG710 = Nothing														'☜: ComPlus Unload
	'		Response.End																'☜: 비지니스 로직 처리를 종료함 
	'	End If  
		
		Set iPAFG710 = Nothing															'☜: Unload Complus

		Response.Write " <Script Language=VBScript> " & vbCr
		Response.Write " Call parent.DbDeleteOk       " & vbCr
		Response.Write " </Script>                  " & vbCr
End Select
%>
