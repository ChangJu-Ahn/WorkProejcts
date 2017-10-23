<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ra3.asp																*
'*  4. Program Name         : Local L/C 상세정보(Local L/C현황조회에서)									*
'*  5. Program Desc         : Local L/C 상세정보(Local L/C현황조회에서)									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/07/12																*
'*  8. Modified date(Last)  : 2002/04/12																*
'*  9. Modifier (First)     : An ChangHwan 																*
'* 10. Modifier (Last)      : Seo Jinkyung															    *
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/07/12 : Coding ReStart											*
'*							  3. 2002/04/12 : ADO 변환													*
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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList													       '☜ : select 대상목록 
Dim lgSelectListDT	

Dim arrRsVal(40)															'☜ : QueryData()실행시 레코드셋을 배열로 받을때 사용	


Dim strMode																		'☜: 현재 MyBiz.asp 의	 진행상태를 나타냄 
Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
    
Call HideStatusWnd
																		'☜ : 받을 레코드셋의 갯수만큼 배열 크기 선언			
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
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
    
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
																		  '조회화면에서 필요한 query조건문들의 영역(Statements table에 있음)
    Redim UNIValue(0,1)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

    UNISqlId(0) = "S3211RA301"  ' main query(spread sheet에 뿌려지는 query statement)

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList         '☜: Select list
																		  '	UNISqlId(0)의 첫번째 ?에 입력됨				
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strVal = ""
	
	strMode = Request("txtMode")														'☜ : 현재 상태를 받음	
	
	if strMode =CStr(UID_M0001) then										
		Err.Clear															
		If Trim(Request("txtLCNo")) = "" Then											
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
			Response.End			
		End If	
	End if
	
	strVal = strVal & " " & filterVar(Request("txtLCNo"),"","S") & " "
	
	if Len(Request("txtSONo")) > 0 then
		strVal= strVal & " And slh.so_no =  " & FilterVar(Trim(Request("txtSONo")), "''", "S") & " "
	End if
	
	UNIValue(0,1) = strVal    '	UNISqlId(0)의 두번째 ?에 입력됨	
    
        
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode 
End Sub

'==========================================================================================================
Sub QueryData()
	Dim iCnt
	Dim FalsechkFlg		
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'☜:ADO 객체를 생성        
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
		'Tab 1 : Local L/C 일반정보 
		
		.txtSONo.value = "<%=ConvSPChars(arrRsVal(1))%>" '수주번호		
		.txtLCDocNo.value = "<%=ConvSPChars(arrRsVal(2))%>"									'
		.txtLCAmendSeq.value = "<%=ConvSPChars(arrRsVal(38))%>"								'
		.txtAdvNo.value = "<%=ConvSPChars(arrRsVal(16))%>"									'통지번호 
		.txtLCType.value = "<%=ConvSPChars(arrRsVal(3))%>"									'
		.txtLCTypeNm.value = "<%=ConvSPChars(arrRsVal(4))%>"									'Local L/C유형 
		.txtAdvDt.text = "<%=UNIDateClientFormat(arrRsVal(17))%>"									'
		.txtFromBank.value = "<%=ConvSPChars(arrRsVal(5))%>"									'추심의뢰은행 
		.txtFromBankNm.value = "<%=ConvSPChars(arrRsVal(6))%>"								'
		.txtExpiryDt.value = "<%=UNIDateClientFormat(arrRsVal(18))%>"								'L/C유효일 
		
		.txtOpenBank.value = "<%=ConvSPChars(arrRsVal(7))%>"									'개설은행 
		.txtOpenBankNm.value = "<%=ConvSPChars(arrRsVal(8))%>"								'
		
		.txtOpenDt.text = "<%=UNIDateClientFormat(arrRsVal(19))%>"

		.txtCurrency.value = "<%=ConvSPChars(arrRsVal(40))%>"							'화폐단위  ?
		parent.CurFormatNumericOCX
		.txtDocAmt.value = "<%=UNIConvNumDBToCompanyByCurrency(arrRsVal(9), arrRsVal(40), ggAmtOfMoneyNo, "X" , "X")%>"	'개설금액 
		.txtLocAmt.value = "<%=UNIConvNumDBToCompanyByCurrency(arrRsVal(20),arrRsVal(40), ggAmtOfMoneyNo, "X" , "X")%>"	'개설자국금액 
		
		.txtXchRate.value = "<%=UNINumClientFormat(arrRsVal(10), ggExchRate.DecPoint, 0)%>"
		.txtRef.value = "<%=ConvSPChars(arrRsVal(21))%>"'선통지사항 
		
		.txtMoveDt.text = "<%=UNIDateClientFormat(arrRsVal(11))%>" '물품인도기일 
		
		
		If "<%=arrRsVal(22)%>" = "Y" Then '분할인도여부 
			.rdoPartailShip1.Checked = True
		ElseIf "<%=arrRsVal(22)%>" = "N" Then
			.rdoPartailShip2.Checked = True
		End If		

		.txtFileDt.value = "<%=ConvSPChars(arrRsVal(12))%>"						'서류제시기간 
		.txtApplicant.value = "<%=ConvSPChars(arrRsVal(23))%>"						'개설신청인 코드 
		.txtApplicantNm.value = "<%=ConvSPChars(arrRsVal(24))%>"						'개설신청인 
		.txtPayTerms.value = "<%=ConvSPChars(arrRsVal(13))%>"							'결제방법코드 
		.txtPayTermsNm.value = "<%=ConvSPChars(arrRsVal(14))%>"						'결제방법 
		.txtBeneficiary.value = "<%=ConvSPChars(arrRsVal(25))%>"						'수혜자 코드 
		.txtBeneficiaryNm.value = "<%=ConvSPChars(arrRsVal(26))%>"					'수혜자						
		.txtAmendDt.text = "<%=UNIDateClientFormat(arrRsVal(15))%>"							'amend일 
		
		.txtSalesGroup.value = "<%=ConvSPChars(arrRsVal(27))%>"						'영업그룹코드 
		.txtSalesGroupNm.value = "<%=ConvSPChars(arrRsVal(28))%>"
		
		'Tab 2 : 구비서류 및 기타 
		
		.txtFileDtTxt.value = "<%=arrRsVal(31)%>"							'서류제시기간 참조 
		.txtDoc1.value = "<%=ConvSPChars(arrRsVal(32))%>"								'구비서류1
		.txtDoc2.value = "<%=ConvSPChars(arrRsVal(33))%>"								'구비서류2	
		.txtDoc3.value = "<%=ConvSPChars(arrRsVal(34))%>"								'구비서류3
		.txtDoc4.value = "<%=ConvSPChars(arrRsVal(35))%>"								'구비서류4
		.txtDoc5.value = "<%=ConvSPChars(arrRsVal(36))%>"								'구비서류5
		.txtBankTxt.value = "<%=ConvSPChars(arrRsVal(37))%>"							'개설은행앞 정보 
		.txtEtcRef.value = "<%=ConvSPChars(arrRsVal(39))%>"							'기타참조사항 
	End With
</Script>	


