<%'======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : Prereceipt
'*  3. Program ID           : f7101mb2
'*  4. Program Name         : ������ ���-���� 
'*  5. Program Desc         : ������ ��� 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2002/11/18
'*  8. Modifier (First)     : ���ͼ� 
'*  9. Modifier (Last)      : Jeong Yong Kyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

														'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True														'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%													                        '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 
Err.Clear 

Call LoadBasisGlobalInf()

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim iPAFG705															    '�� : ����� ComPlus Dll ��� ���� 
Dim IntRows
Dim IntCols
Dim lgIntFlgMode
Dim LngMaxRow
dim iPrrcpt_No
Dim iArrPrrcpt
Dim iArrDept
dim iStrBizPartner
dim iStrCurrency
Dim iStrTransType
Dim iCommandSent
Dim istrTaxBizArea
Dim iErrorPosition

Const C_PRRCPT_NO = 0
Const C_PRRCPT_DT = 1
Const C_REF_NO = 2
Const C_DOC_CUR = 3
Const C_XCH_RATE = 4
Const C_PRRCPT_AMT = 5
Const C_LOC_PRRCPT_AMT = 6
Const C_PRRCPT_STS = 7
Const C_CONF_FG = 8
Const C_PRRCPT_FG = 9
Const C_PRRCPT_DESC = 10
Const C_PRRCPT_TYPE = 11
Const C_VAT_TYPE = 12
Const C_VAT_AMT = 13
Const C_VAT_LOC_AMT = 14
Const C_ISSUED_DT = 15
Const C_PROJECT_NO = 16
Const C_LIMIT_FG = 17

Const C_CHANGEORGID = 0
Const C_DEPT_CD = 1

	' -- ����� 
	' -- ���Ѱ����߰� 
	Const A750_I1_a_data_auth_data_BizAreaCd = 0
	Const A750_I1_a_data_auth_data_internal_cd = 1
	Const A750_I1_a_data_auth_data_sub_internal_cd = 2
	Const A750_I1_a_data_auth_data_auth_usr_id = 3

	Dim I1_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	LngMaxRow = CInt(Request("txtMaxRows"))											'��: �ִ� ������Ʈ�� ���� 

	lgIntFlgMode = CInt(Request("txtFlgMode"))										'��: ����� Create/Update �Ǻ� 

	Set iPAFG705 = Server.CreateObject("PAFG705.cFMngPrSvr")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then					
	   Response.End 
	End If    

	'-----------------------
	'Data manipulate area
	'-----------------------														'��: Single ����Ÿ ���� 
	Redim iArrPrrcpt(C_LIMIT_FG)
	Redim iArrDept(1)

	iStrTransType	             = "FR001"
	iArrPrrcpt(C_PRRCPT_NO)      = Trim(Request("txtPrrcptNo"))
	iArrPrrcpt(C_PRRCPT_DT)      = UNIConvDate(Request("txtPrrcptDt"))
	iArrPrrcpt(C_REF_NO)         = Trim(Request("txtRefNo"))
	iArrPrrcpt(C_DOC_CUR)        = UCase(Trim(Request("txtDocCur")))
	iArrPrrcpt(C_XCH_RATE)       = UNIConvNum(Request("txtXchRate"),0)
	iArrPrrcpt(C_PRRCPT_AMT)     = UNIConvNum(Request("txtPrrcptAmt"),0)
	iArrPrrcpt(C_LOC_PRRCPT_AMT) = ""
	iArrPrrcpt(C_PRRCPT_STS)     = ""
	iArrPrrcpt(C_CONF_FG)        = ""
	iArrPrrcpt(C_PRRCPT_FG)      = "PC"
	iArrPrrcpt(C_PRRCPT_DESC)    = Trim(Request("txtPrrcptDesc"))
	iArrPrrcpt(C_PRRCPT_TYPE)    = UCase(Trim(Request("txtPrrcptType")))
	iArrPrrcpt(C_VAT_TYPE)       = UCase(Trim(Request("txtVatType")))
	iArrPrrcpt(C_VAT_AMT)        = UNIConvNum(Request("txtVatAmt"),0)
	iArrPrrcpt(C_VAT_LOC_AMT)    = UNIConvNum(Request("txtVatLocAmt"),0)
	iArrPrrcpt(C_PROJECT_NO)	 = UCase(Trim(Request("txtProjectNo")))
	iArrPrrcpt(C_LIMIT_FG)		 = Trim(Request("txtLimitFg"))
		
	If Len(Trim(Request("txtVatType"))) > 0 then                                'if VAT is occured(when update)
		iArrPrrcpt(C_VAT_LOC_AMT) = UNIConvNum(Request("txtVatLocAmt"),0)
	Else
		iArrPrrcpt(C_VAT_LOC_AMT) = 0	                                    
	End If		
	If Trim(Request("txtIssuedDt")) = "" Or isnull(Trim(Request("txtIssuedDt"))) Then
		iArrPrrcpt(C_ISSUED_DT)  = UNIConvDate(Request("txtPrrcptDt"))
	Else
		iArrPrrcpt(C_ISSUED_DT)  = UNIConvDate(Request("txtIssuedDt"))
	End if		

	iStrBizPartner               = UCase(Trim(Request("txtBpCd")))

	iArrdept(C_DEPT_CD)          = UCase(Trim(Request("txtDeptCd")))
	iArrDept(C_CHANGEORGID)      = Trim(request("hOrgChangeId"))

	iStrCurrency                 = gCurrency
	iStrTaxBizArea               = UCase(Trim(Request("txtBizAreaCD")))
	

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	iPrrcptNo = iPAFG705.F_MANAGE_PRRCPT_SVR(gStrGloBalCollection,iCommandSent,iStrTransType, _
	                                         iStrCurrency,iArrDept,iStrBizPartner,iArrPrrcpt, _
	                                         iStrTaxBizArea,Trim(Request("txtSpread")),iErrorPosition,I1_a_data_auth)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG705 = Nothing
		Response.End 
	End If    
	
'    if CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then					
'	    Set iPAFG705 = Nothing
'	    Response.End 
'	End If    

	Response.Write " <Script Language=vbscript>                                     " & vbCr
	Response.Write " With parent                                                    " & vbCr   '��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write " .frm1.txtPrrcptNo.Value	= """ & ConvSPChars(iPrrcptNo) & """" & vbCr
	Response.Write " .DbSaveOk                                                      " & vbCr
	Response.Write "  End With                                                      " & vbCr
	Response.Write " </Script>                                                      " & vbCr

	Set iPAFG705 = Nothing															           '��: Unload Complus
%>
