
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : A5406mb1
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
Dim pCardClsAcct																'�� : ��ȸ�� ComProxy Dll ��� ���� 

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

' Com+ Conv. ���� ���� 
Dim pvStrGlobalCollection 


Dim I1_cls_no
Dim E1_a_open_acct
Dim E2_a_cls_acct
Dim E3_a_gl_desc
Dim E4_b_acct_dept


Dim arrCount
Dim lgCurrency

'[CONVERSION INFORMATION]  View Name : export a_open_acct, a_cls_acct_item

Const E1_mgnt_val1 = 0
Const E1_mgnt_val2 = 1
Const E1_gl_no = 2
Const E1_user = 3
Const E1_bank_nm = 4
Const E1_bank_acct = 5
Const E1_amt = 6
Const E1_desc = 7
'[CONVERSION INFORMATION]  View Name : export a_cls_acct
Const E2_cls_dt = 0
Const E2_cls_temp_gl_no = 1
Const E2_cls_gl_no = 2

'[CONVERSION INFORMATION]  View Name : export a_gl
Const E3_gl_desc = 0

'[CONVERSION INFORMATION]  View Name : export a_gl
Const E4_dept_cd = 0
Const E4_dept_nm = 1

				
					'��: ���� ��ȸ/Prev/Next ��û�� ���� 
	'#########################################################################################################
	'												2.2. ��û ���� ó�� 
	'##########################################################################################################
	lgStrPrevKey = Request("lgStrPrevKey")

	'#########################################################################################################
	'												2.3. ���� ó�� 
	'##########################################################################################################

	Set pCardClsAcct = Server.CreateObject("PAUG035.cALkUpCardClsAcctSvr")
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pCardClsAcct = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)	'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

		LngMaxRow  = CLng(Request("txtMaxRows"))												'��: Fetechd Count      
		LngMaxRow1  = CLng(Request("txtMaxRows1"))

		I1_cls_no = Request("txtClsNo")
		On Error Resume next
		Call pCardClsAcct.A_LOOKUP_CARD_CLS_ACCT_SVR(gStrGlobalCollection,Trim(I1_cls_no),E1_a_card_acct,E2_a_cls_acct,E3_a_gl_desc,E4_b_acct_dept)

	'-----------------------
	'Com Action Area
	'-----------------------

		If CheckSYSTEMError(Err,True) = True Then
		
			Set pCardClsAcct = Nothing																	'��: ComProxy Unload
			Response.End																			'��: �����Ͻ� ���� ó���� ������ 
		End If

		Set pCardClsAcct = Nothing
		

		
		Response.Write "<Script Language=vbscript>" & vbCr

		Response.Write " With parent " & vbCr


		Response.Write ".frm1.txtClsDt.text			= """ & UNIDateClientFormat(E2_a_cls_acct(E2_cls_dt))	& """" & vbCr
		Response.Write ".frm1.txtTempGlNo.Value		= """ & ConvSPChars(E2_a_cls_acct(E2_cls_temp_gl_no))	& """" & vbCr
		Response.Write ".frm1.txtGlNo.Value			= """ & ConvSPChars(E2_a_cls_acct(E2_cls_gl_no))	& """" & vbCr

		Response.Write ".frm1.txtDeptCd.Value		= """ & ConvSPChars(E4_b_acct_dept(E4_dept_cd))	& """" & vbCr
		Response.Write ".frm1.txtDeptNm.Value		= """ & ConvSPChars(E4_b_acct_dept(E4_dept_nm))	& """" & vbCr
		Response.Write ".frm1.txtGlDesc.Value		= """ & ConvSPChars(E3_a_gl_desc)	& """" & vbCr
			
		Response.Write " End With "                 & vbCr
	    Response.Write "</Script>"	

	    
	    intCount = UBound(E1_a_card_acct,1)
		StrNextKey = ""   ' import view

		If IsEmpty(E1_a_card_acct) = False and IsArray(E1_a_card_acct) = True Then    

		For IntRows = 0 To intCount		

			strData = strData & Chr(11) & ConvSPChars(E1_a_card_acct(IntRows,E1_mgnt_val1))
			strData = strData & Chr(11) & ConvSPChars(E1_a_card_acct(IntRows,E1_mgnt_val2))
			strData = strData & Chr(11) & ConvSPChars(E1_a_card_acct(IntRows,E1_gl_no))
			strData = strData & Chr(11) & ConvSPChars(E1_a_card_acct(IntRows,E1_user))
			strData = strData & Chr(11) & ConvSPChars(E1_a_card_acct(IntRows,E1_bank_nm))
			strData = strData & Chr(11) & ConvSPChars(E1_a_card_acct(IntRows,E1_bank_acct))
			strData = strData & Chr(11) & UNIConvNumDBToCompanyByCurrency(E1_a_card_acct(IntRows,E1_amt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")
			strData = strData & Chr(11) & ConvSPChars(E1_a_card_acct(IntRows,E1_desc))
	        strData = strData & Chr(11) & LngMaxRow + IntRows
			strData = strData & Chr(11) & Chr(12)                                    
		Next

		End IF

	    Response.Write "<Script Language=VBScript> "                                                          & vbCr  
	    Response.Write " With parent "                                                                        & vbCr 
	    Response.Write " .ggoSpread.Source          = .frm1.vspdData2 "								          & vbCr
	    Response.Write " .ggoSpread.SSShowData        """ & strData											& """" & vbCr
	'    Response.Write " .frm1.vspdData1.MaxRows		= """ & LngMaxRow +  intCount							& """" & vbCr
	    Response.Write " .lgStrPrevKey				= """ & StrNextKey										& """" & vbCr
	    Response.Write " End With "                                                                           & vbCr
	    Response.Write " Call Parent.DbQueryOk2() "                                                                           & vbCr
	    Response.Write "</Script>"  																		  & vbCr	
		
	%>		

