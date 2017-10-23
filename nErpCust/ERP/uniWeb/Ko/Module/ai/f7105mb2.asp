<%'======================================================================================================
'*  1. Module Name          : Finance
'*  2. Function Name        : Prereceipt
'*  3. Program ID           : f7105mb2
'*  4. Program Name         : ������ ���� ���/���� 
'*  5. Program Desc         : ������ ���� ���/���� 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2002/06/19
'*  8. Modifier (First)     : ���ͼ� 
'*  9. Modifier (Last)      : ����� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

Response.Expires = -1														'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True														'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 
Err.Clear 

Call LoadBasisGlobalInf()
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim iPAFG705																'�� : ����� ComPlus Dll ��� ���� 
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
Const C_Gl_Dt = 18

Const C_CHANGEORGID = 0
Const C_DEPT_CD = 1

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
	Redim iArrPrrcpt(C_Gl_Dt)
	Redim iArrDept(1)

	iStrTransType	             = "FR003"
	iArrPrrcpt(C_PRRCPT_NO)      = Trim(Request("txtPrrcptNo"))
	iArrPrrcpt(C_PRRCPT_DT)      = UNIConvDate(Request("txtPrrcptDt"))
	iArrPrrcpt(C_REF_NO)         = Trim(Request("txtRefNo"))
	iArrPrrcpt(C_DOC_CUR)        = UCase(Trim(Request("txtDocCur")))
	iArrPrrcpt(C_XCH_RATE)       = UNIConvNum(Request("txtXchRate"),0)
	iArrPrrcpt(C_PRRCPT_AMT)     = UNIConvNum(Request("txtPrrcptAmt"),0)
	iArrPrrcpt(C_LOC_PRRCPT_AMT) = UNIConvNum(Request("txtPrrcptLocAmt"),0)
	iArrPrrcpt(C_PRRCPT_STS)     = ""
	iArrPrrcpt(C_CONF_FG)        = ""
	iArrPrrcpt(C_PRRCPT_FG)      = "CT"
	iArrPrrcpt(C_PRRCPT_DESC)    = Trim(Request("txtPrrcptDesc"))
	iArrPrrcpt(C_PRRCPT_TYPE)    = UCase(Trim(Request("txtPrrcptType")))
	iArrPrrcpt(C_PROJECT_NO)	 = UCase(Trim(Request("txtProjectNo")))
	iArrPrrcpt(C_LIMIT_FG)		 = Trim(Request("txtLimitFg"))	
	iArrPrrcpt(C_Gl_Dt)			 = UNIConvDate(Request("txtGlDt"))
	
	iStrBizPartner               = UCase(Trim(Request("txtBpCd")))

	iArrdept(C_DEPT_CD)          = UCase(Trim(Request("txtDeptCd")))
	iArrDept(C_CHANGEORGID)      = UCase(Trim(Request("hOrgChangeId")))   'GetGlobalInf("gChangeOrgId")
		
	
	iStrCurrency                 = gCurrency
	iStrTaxBizArea               = Trim(Request("txtBizAreaCD"))

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	iPrrcptNo = iPAFG705.F_MANAGE_PRRCPT_SVR(gStrGloBalCollection,iCommandSent,iStrTransType,iStrCurrency, _
	                                     iArrDept,iStrBizPartner,iArrPrrcpt,Trim(Request("txtSpread")))

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then					
	    Set iPAFG705 = Nothing
	    Response.End 
	End If    

    Response.Write " <Script Language=vbscript>                                  " & vbCr
	Response.Write " With parent                                                 " & vbCr	'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write " .frm1.txtPrrcptNo.Value = """ & ConvSPChars(iPrrcptNo) & """" & vbCr
	Response.Write " .DbSaveOk                                                   " & vbCr
	Response.Write " End With                                                    " & vbCr
	Response.Write " </Script>                                                   " & vbCr

	Set iPAFG705 = Nothing																	'��: Unload Complus
%>
