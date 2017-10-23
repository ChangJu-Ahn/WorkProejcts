<%
Option Explicit 
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a3104mb2
'*  4. Program Name         : �����ݳ������� 
'*  5. Program Desc         : �����ݳ����� ���,���� 
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : ������ 
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************



'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	On Error Resume Next														'��: 

	Call HideStatusWnd()														'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	Call LoadBasisGlobalInf()	
		
	Dim lgIntFlgMode
	Dim LngMaxRow

	Dim iPARG020
	Dim iArrSpread
	Dim ImportTransType
	Dim ImportPartnerBpCd
	Dim ImportDeptCd
	Dim ImportARcpt
	Dim iChangeOrgId

	Dim ExportAGl
	Dim ExportRcpt

	Const RcptNo				= 0
	Const RcptDt				= 1
	Const DocCur				= 2
	Const XchRate				= 3
	Const BnkChgAmt				= 4
	Const BnkChgLocAmt			= 5
	Const RcptType				= 6
	Const RcptDesc				= 7
	Const RefNo					= 8

	'//����ġ�� ����FLAG
	Const Rcpt_Input_Type		= 9		'//����ġ���� 
	Const GlFlag				= 10	'//��ǥ���� 
	Const RcptAmt				= 11
	Const RcptLocAmt			= 12
	Const ConfFg				= 13
	Const Project				= 14

	ReDim ImportARcpt(Project)
	Redim ImportDeptCd(1)

	' -- ���Ѱ����߰� 
	Const A114_I11_a_data_auth_data_BizAreaCd = 0
	Const A114_I11_a_data_auth_data_internal_cd = 1
	Const A114_I11_a_data_auth_data_sub_internal_cd = 2
	Const A114_I11_a_data_auth_data_auth_usr_id = 3

	Dim I11_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

	Redim I11_a_data_auth(3)
	I11_a_data_auth(A114_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
	
	iChangeOrgId = UCase(Request("hOrgChangeId"))

	LngMaxRow					= CInt(Request("txtMaxRows"))					'��: �ִ� ������Ʈ�� ���� 
	lgIntFlgMode				= CInt(Request("txtFlgMode"))					'��: ����� Create/Update �Ǻ� 

	ImportTransType				= "AR001"
	ImportPartnerBpCd			= UCase(Trim(Request("txtBpCd")))	
	
	ImportDeptCd(0)				= iChangeOrgId
	ImportDeptCd(1)				= UCase(Trim(Request("txtDept")))

	ImportARcpt(RcptNo)			= Trim(Request("txtRcptNo"))
	ImportARcpt(RcptDt)			= UNIConvDate(Request("txtRcptDt"))
	ImportARcpt(DocCur)			= UCase(Trim(Request("txtDocCur")))
	ImportARcpt(XchRate)		= UNIConvNum(Request("txtXchRate"),0)
	ImportARcpt(BnkChgAmt)   	= UNIConvNum(Request("txtBankAmt"),0)
	ImportARcpt(BnkChgLocAmt)	= UNIConvNum(Request("txtBankLocAmt"),0)

	If "" & Trim(Request("txtRcptType")) = "" Then
		ImportARcpt(RcptType)	= "H2"
	Else
		ImportARcpt(RcptType)	= UCase(Trim(Request("txtRcptType")))
	End If

	ImportARcpt(RefNo)			= UCase(Trim(Request("txtRefNo")))
	ImportARcpt(RcptDesc)		= Request("txtDesc")
	
	ImportARcpt(Rcpt_Input_Type)= "RP"		'//�����ݵ�� 
	ImportARcpt(ConfFg)			= "U"
	ImportARcpt(Project)		= Trim(Request("txtProject"))
	iArrSpread = Request("txtSpread")
	
	Set iPARG020 = Server.CreateObject("PARG020.cAMngRcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	
	If lgIntFlgMode = OPMD_CMODE Then
		Call iPARG020.MANAGE_RCPT_SVR(gStrGlobalCollection, "CREATE", ImportTransType, gCurrency, _
		                                       ImportARcpt, ImportPartnerBpCd, ImportDeptCd, iArrSpread, _
		                                       ExportAGl, ExportRcpt, I11_a_data_auth)
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		Call  iPARG020.MANAGE_RCPT_SVR(gStrGlobalCollection, "UPDATE", ImportTransType, gCurrency, _
		                                       ImportARcpt, ImportPartnerBpCd, ImportDeptCd, iArrSpread, _
		                                       ExportAGl, ExportRcpt, I11_a_data_auth)
	End If

	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG020 = Nothing		
		Response.End 
	End If
	    
	Set iPARG020 = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.frm1.txtRcptNo.Value	= """ & ConvSPChars(ExportAGl)	& """" & vbcr
	Response.Write " parent.dbSaveOk          " & vbcr
	Response.Write "</Script>" & vbcr
%>
