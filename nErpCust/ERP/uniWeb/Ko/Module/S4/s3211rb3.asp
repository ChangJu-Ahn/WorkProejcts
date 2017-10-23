<%
'********************************************************************************************************
'*  1. Module Name          : ��������																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ra3.asp																*
'*  4. Program Name         : Local L/C ������(Local L/C��Ȳ��ȸ����)									*
'*  5. Program Desc         : Local L/C ������(Local L/C��Ȳ��ȸ����)									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/07/12																*
'*  8. Modified date(Last)  : 2002/04/12																*
'*  9. Modifier (First)     : An ChangHwan 																*
'* 10. Modifier (Last)      : Seo Jinkyung															    *
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : ȭ�� design												*
'*							  2. 2000/07/12 : Coding ReStart											*
'*							  3. 2002/04/12 : ADO ��ȯ													*
'*																										*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "RB")   
Call LoadBNumericFormatB("Q","S","NOCOOKIE","RB")

On Error Resume Next   
Dim iStr		
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList													       '�� : select ����� 
Dim lgSelectListDT	

Dim arrRsVal(40)															'�� : QueryData()����� ���ڵ���� �迭�� ������ ���	


Dim strMode																		'��: ���� MyBiz.asp ��	 ������¸� ��Ÿ�� 
Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
    
Call HideStatusWnd
																		'�� : ���� ���ڵ���� ������ŭ �迭 ũ�� ����			
'--------------- ������ coding part(��������,End)----------------------------------------------------------
lgSelectList= ""

lgSelectList = lgSelectList & " slh.LC_NO  , slh.SO_NO , slh.LC_DOC_NO , slh.L_LC_TYPE "
lgSelectList = lgSelectList & ", lct.minor_nm as lc_type_nm , slh.ADVISE_BANK_CD , ab.bank_nm as advbank , slh.ISSUE_BANK_CD "
lgSelectList = lgSelectList & ", ib.bank_nm as issbank , slh.LC_AMT , slh.XCH_RATE , slh.LATEST_SHIP_DT "
lgSelectList = lgSelectList & ", slh.FILE_DT ,slh.PAY_METH , paym.minor_nm as pay_meth_nm , slh.AMEND_DT "
lgSelectList = lgSelectList & ", slh.ADV_NO , slh.ADV_DT  , slh.EXPIRY_DT, slh.OPEN_DT "

lgSelectList = lgSelectList & ", slh.LC_LOC_AMT , slh.PRE_ADV_REF,  slh.PARTIAL_SHIP_FLAG , slh.APPLICANT "
lgSelectList = lgSelectList & ", ap.bp_nm as app_nm ,  slh.BENEFICIARY , be.bp_nm as be_nm , slh.SALES_GRP "
lgSelectList = lgSelectList & ", sgr.sales_grp_nm, slh.SALES_ORG  , sor.sales_org_nm ,slh.FILE_DT_TXT "
lgSelectList = lgSelectList & ", slh.DOC1 , slh.DOC2 , slh.DOC3 , slh.DOC4 "
lgSelectList = lgSelectList & ", slh.DOC5 , slh.OPEN_BANK_TXT,slh.lc_amend_seq,slh.Remark,slh.cur "



Call FixUNISQLData()
Call QueryData()

'==========================================================================================================
Sub FixUNISQLData()	
    
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
																		  '��ȸȭ�鿡�� �ʿ��� query���ǹ����� ����(Statements table�� ����)
    Redim UNIValue(0,1)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 

    UNISqlId(0) = "S3211RA301"  ' main query(spread sheet�� �ѷ����� query statement)

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList         '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	strVal = ""
	
	strMode = Request("txtMode")														'�� : ���� ���¸� ����	
	
	if strMode =CStr(UID_M0001) then										
		Err.Clear															
		If Trim(Request("txtLCNo")) = "" Then											
			Call ServerMesgBox("��ȸ ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)
			Response.End			
		End If	
	End if
	
	strVal = strVal & " " & filterVar(Request("txtLCNo"),"","S") & " "
	
	if Len(Request("txtSONo")) > 0 then
		strVal= strVal & " And slh.so_no =  " & FilterVar(Trim(Request("txtSONo")), "''", "S") & " "
	End if
	
	UNIValue(0,1) = strVal    '	UNISqlId(0)�� �ι�° ?�� �Էµ�	
    
        
    '--------------- ������ coding part(�������,End)------------------------------------------------------    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode 
End Sub

'==========================================================================================================
Sub QueryData()
	Dim iCnt
	Dim FalsechkFlg		
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'��:ADO ��ü�� ����        
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)    
    FalsechkFlg = False
    
    If  rs0.EOF And rs0.BOF  Then
	
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
    Else
		
		rs0.MoveFirst
		iCnt =0
	
		For iCnt=0 to 40
			arrRsVal(iCnt)=  rs0(iCnt)
		Next
		
        rs0.Close
        Set rs0 = Nothing
        exit sub
    End If   
    

End Sub
%>

<Script Language=vbscript>   	
   	
   	With parent.frm1		
		'Tab 1 : Local L/C �Ϲ����� 
		
		.txtSONo.value = "<%=ConvSPChars(arrRsVal(1))%>" '���ֹ�ȣ		
		.txtLCDocNo.value = "<%=ConvSPChars(arrRsVal(2))%>"									'
		.txtLCAmendSeq.value = "<%=ConvSPChars(arrRsVal(38))%>"								'
		.txtAdvNo.value = "<%=ConvSPChars(arrRsVal(16))%>"									'������ȣ 
		.txtLCType.value = "<%=ConvSPChars(arrRsVal(3))%>"									'
		.txtLCTypeNm.value = "<%=ConvSPChars(arrRsVal(4))%>"									'Local L/C���� 
		.txtAdvDt.text = "<%=UNIDateClientFormat(arrRsVal(17))%>"									'
		.txtFromBank.value = "<%=ConvSPChars(arrRsVal(5))%>"									'�߽��Ƿ����� 
		.txtFromBankNm.value = "<%=ConvSPChars(arrRsVal(6))%>"								'
		.txtExpiryDt.value = "<%=UNIDateClientFormat(arrRsVal(18))%>"								'L/C��ȿ�� 
		
		.txtOpenBank.value = "<%=ConvSPChars(arrRsVal(7))%>"									'�������� 
		.txtOpenBankNm.value = "<%=ConvSPChars(arrRsVal(8))%>"								'
		
		.txtOpenDt.text = "<%=UNIDateClientFormat(arrRsVal(19))%>"

		.txtCurrency.value = "<%=ConvSPChars(arrRsVal(40))%>"							'ȭ�����  ?
		parent.CurFormatNumericOCX
		.txtDocAmt.value = "<%=UNIConvNumDBToCompanyByCurrency(arrRsVal(9), arrRsVal(40), ggAmtOfMoneyNo, "X" , "X")%>"	'�����ݾ� 
		.txtLocAmt.value = "<%=UNIConvNumDBToCompanyByCurrency(arrRsVal(20),arrRsVal(40), ggAmtOfMoneyNo, "X" , "X")%>"	'�����ڱ��ݾ� 
		
		.txtXchRate.value = "<%=UNINumClientFormat(arrRsVal(10), ggExchRate.DecPoint, 0)%>"
		.txtRef.value = "<%=ConvSPChars(arrRsVal(21))%>"'���������� 
		
		.txtMoveDt.text = "<%=UNIDateClientFormat(arrRsVal(11))%>" '��ǰ�ε����� 
		
		
		If "<%=arrRsVal(22)%>" = "Y" Then '�����ε����� 
			.rdoPartailShip1.Checked = True
		ElseIf "<%=arrRsVal(22)%>" = "N" Then
			.rdoPartailShip2.Checked = True
		End If		

		.txtFileDt.value = "<%=ConvSPChars(arrRsVal(12))%>"						'�������ñⰣ 
		.txtApplicant.value = "<%=ConvSPChars(arrRsVal(23))%>"						'������û�� �ڵ� 
		.txtApplicantNm.value = "<%=ConvSPChars(arrRsVal(24))%>"						'������û�� 
		.txtPayTerms.value = "<%=ConvSPChars(arrRsVal(13))%>"							'��������ڵ� 
		.txtPayTermsNm.value = "<%=ConvSPChars(arrRsVal(14))%>"						'������� 
		.txtBeneficiary.value = "<%=ConvSPChars(arrRsVal(25))%>"						'������ �ڵ� 
		.txtBeneficiaryNm.value = "<%=ConvSPChars(arrRsVal(26))%>"					'������						
		.txtAmendDt.text = "<%=UNIDateClientFormat(arrRsVal(15))%>"							'amend�� 
		
		.txtSalesGroup.value = "<%=ConvSPChars(arrRsVal(27))%>"						'�����׷��ڵ� 
		.txtSalesGroupNm.value = "<%=ConvSPChars(arrRsVal(28))%>"
		
		'Tab 2 : ���񼭷� �� ��Ÿ 
		
		.txtFileDtTxt.value = "<%=arrRsVal(31)%>"							'�������ñⰣ ���� 
		.txtDoc1.value = "<%=ConvSPChars(arrRsVal(32))%>"								'���񼭷�1
		.txtDoc2.value = "<%=ConvSPChars(arrRsVal(33))%>"								'���񼭷�2	
		.txtDoc3.value = "<%=ConvSPChars(arrRsVal(34))%>"								'���񼭷�3
		.txtDoc4.value = "<%=ConvSPChars(arrRsVal(35))%>"								'���񼭷�4
		.txtDoc5.value = "<%=ConvSPChars(arrRsVal(36))%>"								'���񼭷�5
		.txtBankTxt.value = "<%=ConvSPChars(arrRsVal(37))%>"							'��������� ���� 
		.txtEtcRef.value = "<%=ConvSPChars(arrRsVal(39))%>"							'��Ÿ�������� 
	End With
</Script>	


