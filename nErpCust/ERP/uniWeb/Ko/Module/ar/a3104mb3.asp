<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a4103mb3
'*  4. Program Name         : �����ݳ������� 
'*  5. Program Desc         : ������ ������ ���� 
'*  6. Complus List         :
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : ���ͼ� 
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
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	On Error Resume Next														'��: 

	Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
	
	Call LoadBasisGlobalInf()

	Dim iPARG020																'�� : ��ȸ�� ComProxy Dll ��� ���� 
	Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
		
	Dim iArrSpread
	Dim ImportTransType
	Dim ImportDocCur
	Dim ImportPartnerBpCd
	Dim ImportDeptCd
	Dim ImportARcpt

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
	Const Gl_Input_Type			= 9		'//����ġ���� 
	Const GlFlag				= 10	'//��ǥ���� 

	ReDim ImportARcpt(GlFlag)

	Redim ImportDeptCd(2)

	' -- ���Ѱ����߰� 
	Const A114_I11_a_data_auth_data_BizAreaCd = 0
	Const A114_I11_a_data_auth_data_internal_cd = 1
	Const A114_I11_a_data_auth_data_sub_internal_cd = 2
	Const A114_I11_a_data_auth_data_auth_usr_id = 3

	Dim I11_a_data_auth  '--> �Ķ������ ������ ���� ���̹� ���� 

	Redim I11_a_data_auth(3)
	I11_a_data_auth(A114_I11_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I11_a_data_auth(A114_I11_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))
	
	strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

	If strMode = "" Then
		Response.End 
	ElseIf strMode <> CStr(UID_M0003) Then										'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End 
	ElseIf Request("txtRcptNo") = "" Then										'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If

	ImportTransType				= "AR001"
	ImportARcpt(RcptNo)			= Trim(Request("txtRcptNo"))

	ImportARcpt(Gl_Input_Type)	= "RP"		'//�����ݵ�� 
	


	Set iPARG020 = Server.CreateObject("PARG020.cAMngRcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	                         
	Call iPARG020.MANAGE_RCPT_SVR(gStrGlobalCollection, "DELETE", ImportTransType, gCurrency, _
		                                       ImportARcpt, ImportPartnerBpCd, ImportDeptCd, iArrSpread, _
		                                       ExportAGl, ExportRcpt, I11_a_data_auth)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG020 = Nothing		
		Response.End 
	End If
	    
	Set iPARG020 = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.DbDeleteOk()      " & vbcr
	Response.Write "</Script>" & vbcr
%>
