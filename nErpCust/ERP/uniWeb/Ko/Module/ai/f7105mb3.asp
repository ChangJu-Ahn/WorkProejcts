<%'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : Prereceipt
'*  3. Program ID           : f7105mb3
'*  4. Program Name         : ������ ���ʵ���� �������� 
'*  5. Program Desc         : ������ ���ʵ���� �������� 
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
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 
Err.Clear 

Call LoadBasisGlobalInf()
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim iPAFG705 																'�� : ��ȸ�� ComPlus Dll ��� ���� 
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim iCommandSent
Dim iTransType 
Dim iarrPrrcpt 

Const A835_I5_prrcpt_no = 0
Const A835_I5_prrcpt_dt = 1
Const A835_I5_ref_no = 2
Const A835_I5_doc_cur = 3
Const A835_I5_xch_rate = 4
Const A835_I5_prrcpt_amt = 5
Const A835_I5_loc_prrcpt_amt = 6
Const A835_I5_prrcpt_sts = 7
Const A835_I5_conf_fg = 8
Const A835_I5_prrcpt_fg = 9
Const A835_I5_prrcpt_desc = 10
Const A835_I5_prrcpt_type = 11
Const A835_I5_vat_type = 12
Const A835_I5_vat_amt = 13
Const A835_I5_vat_loc_amt = 14
Const A835_I5_issued_dt = 15
Const A835_I5_c_limit_fg = 16

	strMode = Request("txtMode")											'�� : ���� ���¸� ���� 

	Set iPAFG705 = Server.CreateObject("PAFG705.cFMngPrSvr")	    

	If CheckSYSTEMError(Err, True) = True Then					
	   Response.End 
	End If    

	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------
	Redim iarrPrrcpt(A835_I5_c_limit_fg)

	iCommandSent = "DELETE"

	iTransType	                  = "FR003"
	iArrPrrcpt(A835_I5_prrcpt_no) = Trim(Request("txtPrrcptNo"))
	iArrPrrcpt(A835_I5_prrcpt_fg) = "CT"

	Call iPAFG705.F_MANAGE_PRRCPT_SVR(gStrGloBalCollection,iCommandSent,iTransType,,,,iarrPrrcpt)

	If CheckSYSTEMError(Err, True) = True Then					
	    Set iPAFG705 = Nothing
	    Response.End 
	End If   

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " Call parent.DbDeleteOk() " & vbCr
	Response.Write "</Script> "                 & vbCr

	Set iPAFG705 = Nothing                                                   '��: Unload Complus
%>
