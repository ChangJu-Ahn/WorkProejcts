<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : REPAY LOAN MULTI DELETE
'*  3. Program ID        : f4255mb3
'*  4. Program �̸�      : ���Աݸ�Ƽ��ȯ(����)
'*  5. Program ����      : ���Աݸ�Ƽ��ȯ 
'*  6. Complus ����Ʈ    : PAFG430.DLL
'*  7. ���� �ۼ������   : 2003/05/10
'*  8. ���� ���������   : 2003/05/10
'*  9. ���� �ۼ���       : ����� 
'* 10. ���� �ۼ���       : ����� 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%																						'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call HideStatusWnd																		'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next																	'��: 
Err.Clear 
 
Call LoadBasisGlobalInf()

Dim iPAFG430																			'�� : ��ȸ�� ComPlus Dll ��� ���� 
Dim strMode																				'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim iCommandSent

Dim I1_f_ln_repay
Const A861_I1_repay_no = 0

strMode = Request("txtMode")															'�� : ���� ���¸� ���� 

	If strMode = "" Then
		Response.End 
		Call HideStatusWnd		
	ElseIf strMode <> CStr(UID_M0003) Then												'�� : ��ȸ ���� Biz �̹Ƿ� �ٸ����� �׳� ������ 
		Response.End 
		Call HideStatusWnd		
	ElseIf Trim(Request("txtRePayNo")) = "" Then											'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)						'��ȸ ���ǰ��� ����ֽ��ϴ�!
		Response.End
		Call HideStatusWnd		 
	End If

	iCommandSent = "DELETE"

	Redim I1_f_ln_repay(6)
	I1_f_ln_repay(A861_I1_repay_no) = Trim(Request("txtRePayNO"))

	Set iPAFG430 = Server.CreateObject("PAFG430.cFMngRepayMultiSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																	'��: �����Ͻ� ���� ó���� ������ 
	End If	

	Call iPAFG430.F_MANAGE_REPAY_MULTI_SVR(gStrGlobalCollection,iCommandSent, I1_f_ln_repay)						
		
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG430 = Nothing															'��: ComProxy Unload
		Response.End																	'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	Set iPAFG430 = Nothing																'��: Unload Complus	
	                                                
	Response.Write " <Script Language=vbscript> " & vbCr
   	Response.Write " Call parent.DbDeleteOk()   " & vbCr
	Response.Write " </Script>                  " & vbCr
%>
