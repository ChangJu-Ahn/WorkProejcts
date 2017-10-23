<%@ LANGUAGE=VBSCript %>
<%Option Explicit
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a3104mb1
'*  4. Program Name         : �����ݳ�����ȸ 
'*  5. Program Desc         : �����ݳ�����ȸ 
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/10/13
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

'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. ���Ǻ� 
'##########################################################################################################
																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
On Error Resume Next														'��: 

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'#########################################################################################################
'												2.1 ���� üũ 
'##########################################################################################################
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then										'��: ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
	Call ServerMesgBox("700118", vbInformation, I_MKSCRIPT)					'��: ��ȸ �����ε� �ٸ� ���·� ��û�� ���� ���, �ʿ������ ���� ��, �޼����� ID������ ����ؾ� �� 
	Response.End 
ElseIf Request("txtRcptNo") = "" Then										'��: ��ȸ�� ���� ���� ���Դ��� üũ 
	Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)					'��:
	Response.End 
End If

'#########################################################################################################
'												2. ���� ó�� ����� 
'##########################################################################################################

'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
Dim iArrData
Dim iGData
Dim lgStrPrevKey
Dim iLngRow
Dim LngMaxRow
Dim iARcptItemSeq
Dim iPARG020
Dim iStrData
Dim lgCurrency
Dim iRcptNo
Dim iRcptInputType

Const RcptNo = 0
Const JnlCd = 1
Const JnlNm = 2
Const ConfFg = 3
Const DeptCd = 4
Const DeptNm = 5
Const RcptDt = 6
Const BpCd = 7
Const BpNm = 8
Const RefNo = 9
Const DocCur = 10
Const XchRate = 11
Const RcptAmt = 12
Const RcptLocAmt = 13
Const BnkChgAmt = 14
Const BnkChgLocAmt = 15
Const AllcAmt = 16
Const AllclocAmt = 17
Const Adjustamt = 18
Const AdjustLocAmt = 19
Const BalAmt = 20
Const BalLocAmt = 21
Const TempGlNo = 22
Const GlNo = 23
Const RcptDesc = 24
Const Project = 25

Const EG1_E1_rcpt_type = 0
Const EG1_E1_rcpt_type_nm = 1           
Const EG1_E1_net_rcpt_amt = 2
Const EG1_E1_net_rcpt_loc_amt = 3
Const EG1_E1_note_no = 4
Const EG1_E1_seq = 5                
Const EG1_E1_bank_acct_no = 6       
Const EG1_E1_acct_cd = 7            
Const EG1_E1_acct_nm = 8
Const EG1_E1_item_desc = 9

' -- ���Ѱ����߰� 
Const A500_I4_a_data_auth_data_BizAreaCd = 0
Const A500_I4_a_data_auth_data_internal_cd = 1
Const A500_I4_a_data_auth_data_sub_internal_cd = 2
Const A500_I4_a_data_auth_data_auth_usr_id = 3

Dim I4_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

Redim I4_a_data_auth(3)
I4_a_data_auth(A500_I4_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
I4_a_data_auth(A500_I4_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
I4_a_data_auth(A500_I4_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
I4_a_data_auth(A500_I4_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

'#########################################################################################################
'												2.2. ��û ���� ó�� 
'##########################################################################################################
	lgStrPrevKey = Request("lgStrPrevKey")
	LngMaxRow = Request("txtMaxRows")
'#########################################################################################################
'												2.3. ���� ó�� 
'##########################################################################################################

	If lgStrPrevKey = "" Then
		iARcptItemSeq = 0
	Else
		iARcptItemSeq = lgStrPrevKey
	End If

	Set iPARG020 = Server.CreateObject("PARG020.cALkUpRcSvr")

	If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
	End If    
		
	iRcptNo = Trim(Request("txtRcptNo"))
	iRcptInputType = "RP"
	Call iPARG020.LOOKUP_RCPT_SVR(gStrGloBalCollection, iARcptItemSeq, iRcptNo, iRcptInputType,iArrData, iGData, I4_a_data_auth)
		
	If CheckSYSTEMError(Err, True) = True Then					
	   Set iPARG020 = Nothing
	   Response.End 
	End If    

	lgCurrency = iArrDAta(DocCur)

	Response.Write "<Script Language=vbscript>  " & vbcr
	Response.Write " With parent.frm1           " & vbcr														'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write ".txtRcptNo.Value		= """ & ConvSPChars(iArrData(RcptNo))			& """ " & vbcr
	Response.Write ".txtRcptType.value		= """ & ConvSPChars(iArrData(JnlCd))			& """ " & vbcr
	Response.Write ".txtRcptTypeNm.value	= """ & ConvSPChars(iArrData(JnlNm))			& """ " & vbcr
	Response.Write ".txtDeptNm.Value		= """ & ConvSPChars(iArrData(DeptNm))			& """ " & vbcr
	Response.Write ".txtDept.Value			= """ & ConvSPChars(iArrData(DeptCd))			& """ " & vbcr
	Response.Write ".fpDateTime1.Text       = """ & UNIDateClientFormat(iArrData(RcptDt))	& """ " & vbcr
	Response.Write ".txtBpCd.Value			= """ & ConvSPChars(iArrData(BpCd))				& """ " & vbcr
	Response.Write ".txtBpNm.Value			= """ & ConvSPChars(iArrData(BpNm))				& """ " & vbcr
	Response.Write ".txtRefNo.value			= """ & ConvSPChars(iArrDAta(RefNo))			& """ " & vbcr
	Response.Write ".txtDocCur.Value		= """ & ConvSPChars(iArrDAta(DocCur))			& """ " & vbcr
	Response.Write ".txtXchRate.Value		= """ & UNINumClientFormat(iArrDAta(XchRate), ggExchRate.DecPoint, 0)			& """ " &vbcr

	Response.Write ".txtRcptAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(RcptAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """ " &vbcr
	Response.Write ".txtRcptLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(RcptLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """ " & vbcr

	Response.Write ".txtBankAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(BnkChgAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """ " & vbcr 
	Response.Write ".txtBankLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(BnkChgLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """ " & vbcr

	Response.Write ".txtClsAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AllcAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """ " & vbcr 
	Response.Write ".txtClsLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AllcLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """ " & vbcr

	Response.Write ".txtSttlAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AdjustAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """ " & vbcr 
	Response.Write ".txtSttlLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AdjustLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """ " & vbcr

	Response.Write ".txtBalAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iArrData(BalAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")			& """ " & vbcr 
	Response.Write ".txtBalLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(BalLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """ " & vbcr

	Response.Write ".txtTempGLNo.Value		= """ & ConvSPChars(iArrData(TempGlNo))			& """ " & vbcr        
	Response.Write ".txtGlNo.Value			= """ & ConvSPChars(iArrData(GlNo))				& """ " & vbcr
	Response.Write ".txtDesc.value			= """ & ConvSPChars(iArrDAta(RcptDesc))			& """ " & vbcr
	Response.Write ".txtProject.value		= """ & ConvSPChars(iArrDAta(Project))			& """ " & vbcr
	Response.Write " End With					" & vbcr		    
	Response.Write " Parent.DbQueryOk			" & vbcr
	Response.write "</Script>				    " & vbcr  

	iStrData = ""
		
	For iLngRow = 0 To UBound(iGData, 1) 	
			iStrData = iStrData & Chr(11) & ConvSPChars(iGData(iLngRow, EG1_E1_rcpt_type))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(iGData(iLngRow, EG1_E1_rcpt_type_nm))	
			iStrData = iStrData & Chr(11) & ConvSPChars(iGData(iLngRow, EG1_E1_acct_cd))
			iStrData = iStrData & Chr(11) & ""			
			iStrData = iStrData & Chr(11) & ConvSPChars(iGData(iLngRow, EG1_E1_acct_nm))
			iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iGData(iLngRow, EG1_E1_net_rcpt_amt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")
			iStrData = iStrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(iGData(iLngRow, EG1_E1_net_rcpt_loc_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		
			iStrData = iStrData & Chr(11) & ConvSPChars(iGData(iLngRow, EG1_E1_note_no))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(iGData(iLngRow, EG1_E1_bank_acct_no))	
			iStrData = iStrData & Chr(11) & ""			
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(iGData(iLngRow, EG1_E1_item_desc))	
			iStrData = iStrData & Chr(11) & ""															
			iStrData = iStrData & Chr(11) & LngMaxRow + iLngRow + 1
			iStrData = iStrData & Chr(11) & Chr(12)
	Next

	Response.Write " <Script Language=vbscript>								" & vbCr
	Response.Write " With parent											" & vbCr
	Response.Write "	.ggoSpread.Source = .frm1.vspdData			" & vbcr
	Response.Write "	.ggoSpread.SSShowData """ & istrData & """	" & vbcr
	Response.Write "	.DbQueryOk()										" & vbcr
    Response.Write " End With												" & vbCr
    Response.Write " </Script>												" & vbCr
%>

