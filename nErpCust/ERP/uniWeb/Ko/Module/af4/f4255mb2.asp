<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : REPAY LOAN MULTI SAVE
'*  3. Program ID        : f4255mb2
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
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																								'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd()																			'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next																			'��: 
Err.Clear 

Call LoadBasisGlobalInf()

Dim iPAFG430																					'����� ComPlus Dll ��� ���� 
Dim lgIntFlgMode
Dim iCommandSent 

Dim I1_f_ln_repay
Const A861_I1_repay_no = 0
Const A861_I1_repay_dt = 1
Const A861_I1_repay_dept_cd = 2
Const A861_I1_repay_org_change_id = 3
Const A861_I1_repay_user_fld1 = 4
Const A861_I1_repay_user_fld2 = 5
Const A861_I1_repay_desc = 6

Dim E1_b_auto_numbering 

	lgIntFlgMode = CInt(Request("txtMode"))														'��: ����� Create/Update �Ǻ� 

	'-----------------------
	'Data manipulate area
	'---------- -------------																	'Single ����Ÿ ���� 
	ReDim I1_f_ln_repay(A861_I1_repay_desc)
	
	I1_f_ln_repay(A861_I1_repay_no) = Trim(Request("txtRePayNO"))
	I1_f_ln_repay(A861_I1_repay_dt) = UNIConvDate(Request("txtRePayDT"))
	I1_f_ln_repay(A861_I1_repay_org_change_id) = UCase(Request("horgChangeId"))
	I1_f_ln_repay(A861_I1_repay_dept_cd) = UCase(Trim(Request("txtDeptCd")))
	I1_f_ln_repay(A861_I1_repay_user_fld1) = Trim(Request("txtUserFld1"))
	I1_f_ln_repay(A861_I1_repay_user_fld2) = Trim(Request("txtUserFld2"))
	I1_f_ln_repay(A861_I1_repay_desc) = Trim(Request("txtRePayDesc"))

	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	'-----------------------------------------
	'Com Action Area
	'-----------------------------------------
	Set iPAFG430 = Server.CreateObject("PAFG430.cFMngRepayMultiSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If		

	Call iPAFG430.F_MANAGE_REPAY_MULTI_SVR(gStrGlobalCollection,iCommandSent, I1_f_ln_repay, _
								Trim(Request("txtSpread4")),Trim(Request("txtSpread1")), _
								Trim(Request("txtSpread")), E1_b_auto_numbering)						

	'---------------------------------------------
	'Com action result check area(OS,internal)
	'---------------------------------------------
	If CheckSYSTEMError(Err,True) = True Then
		Set iPAFG430 = Nothing																	'��: ComPlus Unload
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If		
		
	Set iPAFG430 = Nothing																		'��: ComPlus Unload		


    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
    Response.Write " .DBSaveOk(""" & ConvSPChars(E1_b_auto_numbering)  & """)" & vbCr
    Response.Write "End With "					 & vbCr	  
    Response.Write "</Script>"           
%>
