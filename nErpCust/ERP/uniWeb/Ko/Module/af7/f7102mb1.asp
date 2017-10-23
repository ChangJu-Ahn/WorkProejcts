
<%'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : PreReceipt
'*  3. Program ID        : f7102mb1
'*  4. Program �̸�      : ������ û�� 
'*  5. Program ����      : ������ û�� ����Ʈ ��ȸ , û���߰� , ���� , ���� 
'*  6. Comproxy ����Ʈ   : fr0021 , fr0028
'*  7. ���� �ۼ������   : 2000/10/7
'*  8. ���� ���������   : 2002/11/19
'*  9. ���� �ۼ���       : �ۺ��� 
'* 10. ���� �ۼ���       : Jeong Yong Kyun
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************
														'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True														'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call HideStatusWnd

On Error Resume Next														'��: 
Err.Clear 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim StrNextKey																' ���� �� 
Dim lgStrPrevKey															' ���� �� 
Dim LngMaxRow																' ���� �׸����� �ִ�Row
Dim LngMaxRow3																' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount 
Dim LngLastRow      
Dim strData
Dim lgCurrency       

Dim iPAFG710																'�Է�/������ ComPlus Dll ��� ���� 
Dim iprrcptprrcptNo
Dim iprrcptSttlMentNo 
Dim iprrcptSttlDt
Dim iprrcptDocCur

Dim arrVal, arrTemp																'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus																	'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim	lGrpCnt																		'��: Group Count
Dim strCode																		'Lookup �� ���� ���� 
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
Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

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
	Dim iarrRpSttl		'//for leu project : û������, �ŷ���ȭ, ��ǥ��ȣ(����, ȸ��)

	' -- ��ȸ�� 
	' -- ���Ѱ����߰� 
	Const A838_I2_a_data_auth_data_BizAreaCd = 0
	Const A838_I2_a_data_auth_data_internal_cd = 1
	Const A838_I2_a_data_auth_data_sub_internal_cd = 2
	Const A838_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

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
		Response.End								'��: �����Ͻ� ���� ó���� ������ 
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
		Set iPAFG710 = Nothing						'��: ComPlus Unload
		Response.End								'��: �����Ͻ� ���� ó���� ������ 
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
    Response.Write " With parent                " & vbCr										'��: ȭ�� ó�� ASP �� ��Ī�� 
	
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
Case CStr(UID_M0002)																'��: ���� ��û�� ���� 

	' -- ����� 
	' -- ���Ѱ����߰� 
	Const A837_I6_a_data_auth_data_BizAreaCd = 0
	Const A837_I6_a_data_auth_data_internal_cd = 1
	Const A837_I6_a_data_auth_data_sub_internal_cd = 2
	Const A837_I6_a_data_auth_data_auth_usr_id = 3

	Dim I6_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I6_a_data_auth(3)
	I6_a_data_auth(A837_I6_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I6_a_data_auth(A837_I6_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I6_a_data_auth(A837_I6_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I6_a_data_auth(A837_I6_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
												
    Err.Clear																		'��: Protect system from crashing

    LngMaxRow  = CInt(Request("txtMaxRows"))										'��: �ִ� ������Ʈ�� ���� 
    LngMaxRow3 = CInt(Request("txtMaxRows3"))

    Set iPAFG710 = Server.CreateObject("PAFG710.cFMngPrSttlSvr")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    
    If CheckSYSTEMError(Err, True) = True Then					
		Response.End																'��: �����Ͻ� ���� ó���� ������ 
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
    	Set iPAFG710 = Nothing														'��: ComPlus Unload
		Response.End																'��: �����Ͻ� ���� ó���� ������ 
	End If   
    
'   if CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then					
'		Set iPAFG710 = Nothing														'��: ComPlus Unload
'		Response.End																'��: �����Ͻ� ���� ó���� ������ 
'	End If  
	
	Set iPAFG710 = Nothing															'��: Unload Complus

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
	Case CStr(UID_M0003)																'��: ���� ��û�� ���� 

		Err.Clear																		'��: Protect system from crashing

	' -- ��ȸ�� 
	' -- ���Ѱ����߰� 
	Const A697_I6_a_data_auth_data_BizAreaCd = 0
	Const A697_I6_a_data_auth_data_internal_cd = 1
	Const A697_I6_a_data_auth_data_sub_internal_cd = 2
	Const A697_I6_a_data_auth_data_auth_usr_id = 3

	'Dim I6_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

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
			Response.End																'��: �����Ͻ� ���� ó���� ������ 
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
	    	Set iPAFG710 = Nothing														'��: ComPlus Unload
			Response.End																'��: �����Ͻ� ���� ó���� ������ 
		End If   
	    
	'   if CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then					
	'		Set iPAFG710 = Nothing														'��: ComPlus Unload
	'		Response.End																'��: �����Ͻ� ���� ó���� ������ 
	'	End If  
		
		Set iPAFG710 = Nothing															'��: Unload Complus

		Response.Write " <Script Language=VBScript> " & vbCr
		Response.Write " Call parent.DbDeleteOk       " & vbCr
		Response.Write " </Script>                  " & vbCr
End Select
%>
