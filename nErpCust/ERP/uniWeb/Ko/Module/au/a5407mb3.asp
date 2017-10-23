
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : A5407mb1
'*  4. Program Name         : �̰����(�������)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2002/11/05
'*  8. Modified date(Last)  : 2002/11/05
'*  9. Modifier (First)     : KIM HO YOUNG
'* 10. Modifier (Last)      : KIM HO YOUNG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2002/11/05 : ..........
'**********************************************************************************************


Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ����Ǿ� �������� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

%>
<%
'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. ���Ǻ� 
'##########################################################################################################

													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd	
On Error Resume Next														'��: 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

'#########################################################################################################
'												2.1 ���� üũ 
'##########################################################################################################

If strMode = "" Then
'	Response.End 
End If

'#########################################################################################################
'												2. ���� ó�� ����� 
'##########################################################################################################

'#########################################################################################################
'												2.1. ����, ��� ���� 
'##########################################################################################################
' ������ ���� 
Dim pGlCardAcct																'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim IntRows
Dim IntCols
Dim sList
Dim vbIntRet
Dim intCount
Dim IntCount1
Dim LngMaxRow
Dim LngMaxRow1
Dim StrNextKey
Dim lgStrPrevKey
Dim lgIntFlgMode
dim test
Dim I1_mgnt_acct_cd
Dim I2_b_bank
Dim I3_trans_type
Dim I4_a_gl
Dim I5_gCurrency
Dim arrCount
Dim lgCurrency
Dim E1_cls_no
' Com+ Conv. ���� ���� 
Dim pvStrGlobalCollection 

Const I2_b_bank_cd = 0
Const I2_b_bank_acct = 1

Const I4_gl_dt = 0
Const I4_gl_desc = 1
Const I4_gl_input_type = 2
Const I4_dept_cd = 3
Const I4_org_change_id = 4

Dim I6_txtFromBaseDt
Dim I7_txtToBaseDt

ReDim I2_b_bank(I2_b_bank_acct)    
ReDim I4_a_gl(I4_org_change_id)

'[CONVERSION INFORMATION]  View Name : export a_open_acct



'#########################################################################################################
'												2.2. ��û ���� ó�� 
'##########################################################################################################
lgStrPrevKey = Request("lgStrPrevKey")

I1_mgnt_acct_cd = Trim(Request("hmgnt_acct_cd")) ' ���¿����� �����ڵ� 

I2_b_bank(I2_b_bank_cd) = Trim(Request("htxtBankCd"))
I2_b_bank(I2_b_bank_acct) = Trim(Request("htxtBankAcct"))

I3_trans_type = "AP011"

I4_a_gl(I4_gl_dt) = UNIConvDate(Request("htxtGlDt"))
I4_a_gl(I4_gl_desc) = Trim(Request("htxtDesc"))
I4_a_gl(I4_gl_input_type) = "OC"
I4_a_gl(I4_dept_cd) = Trim(Request("htxtDeptCd"))
I4_a_gl(I4_org_change_id) = Trim(Request("htxtOrgChangeId"))

I5_gCurrency = gCurrency


'#########################################################################################################
'												2.3. ���� ó�� 
'##########################################################################################################


Set pGlCardAcct = Server.CreateObject("PAUG035.cACreGlCardAcctSvr")
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then
	Set pOpenCardAcct = Nothing												'��: ComProxy Unload
	Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)	'��:
	Response.End														'��: �����Ͻ� ���� ó���� ������ 
End If

	LngMaxRow  = CLng(Request("txtMaxRows"))												'��: Fetechd Count      
	LngMaxRow1  = CLng(Request("txtMaxRows1"))

	I6_txtFromBaseDt = UniConvDateAToB(Trim(Request("txtFromBaseDt")),gDateFormat, gServerDateFormat)
	I7_txtToBaseDt = UniConvDateAToB(Trim(Request("txtToBaseDt")),gDateFormat, gServerDateFormat)


	Call pGlCardAcct.A_CREATE_GL_CARD_ACCT_SVR(gStrGlobalCollection,I1_mgnt_acct_cd,I2_b_bank,I3_trans_type,_
						I4_a_gl,I5_gCurrency,I6_txtFromBaseDt, I7_txtToBaseDt, E1_cls_no)
'-----------------------
'Com Action Area
'-----------------------

	If CheckSYSTEMError(Err,True) = True Then
	
		Set pGlCardAcct = Nothing																	'��: ComProxy Unload
		Response.End																			'��: �����Ͻ� ���� ó���� ������ 
	End If

	Set pGlCardAcct = Nothing
    
    Response.Write "<Script Language=VBScript> " & vbCr         
    Response.Write "With parent "				 & vbCr	
	Response.Write " .frm1.txtClsNo.value = """ & ConvSPChars(Trim(E1_cls_no)) & """" & vbCr
	Response.Write " .DbSaveOK() "								  & vbCr
    Response.Write "End With "				 & vbCr	  
    Response.Write "</Script>"     	
%>		

