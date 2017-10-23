<%
Option Explicit
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 기초 Open Ap 조회하는 p/g
'*  3. Program ID           : a4110mb1	
'*  4. Program Name         :	
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/14
'*  9. Modifier (First)     : YOU SO EUN
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되어 마지막에 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2.- 조건부 
'##########################################################################################################
																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
'On Error Resume Next														'☜: 
'Err.Clear 

Call HideStatusWnd()														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'#########################################################################################################
'												2.1 조건 체크 
'##########################################################################################################
If strMode = "" Then'
	Response.End
	Call HideStatusWnd()		 		 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
Dim iPAPG005																'☆ : 조회용 ComPlus Dll 사용 변수 
Dim ImportData 
Dim ExportData1
Dim ExportData2
Dim lgCurrency

Const acctcd		= 0
Const acctnm		= 1
Const BpCd			= 2
Const BpNm			= 3
Const apno			= 4
Const apdt			= 5
Const invdocno		= 6
Const invdt			= 7
Const BlDocNo		= 8
Const bldt			= 9
Const refno			= 10
Const doccur		= 11
Const xchrate		= 12
Const DueDt			= 13
Const NetAmt		= 14
Const Netlocamt		= 15
Const VatAmt		= 16
Const VatLocAmt		= 17
Const ApAmt			= 18
Const ApLocAmt		= 19
Const ApType		= 20
Const PaymType		= 21
Const Paymterms		= 22
Const ApDesc		= 23
Const LcDocNo		= 24
Const ApLcDt		= 25
Const CashAmt		= 26
Const CashLocAmt	= 27
Const PrPaymamt		= 28
Const PrPaymlocamt	= 29
Const PrPaymno		= 30
Const BalAmt		= 31
Const BalLocAmt		= 32
Const TotApAmt		= 33
Const TotApLocAmt	= 34
Const PayMeth		= 35
Const PayDur		= 36
Const tempglno		= 37
Const DeptCd		= 38
Const DeptNm		= 39
Const ReportBpCd	= 40
Const ReportBpNm	= 41
Const PayBpCd		= 42
Const PayBpNm		= 43
Const bizareacd		= 44
Const bizareanm		= 45
Const PaymTypeNm	= 46
Const PayMethNm		= 47
Const ConfFg		= 48
Const GlNo			= 49
Const Gldt			= 50

Const gIsShowLocal = "Y"

'#########################################################################################################
'												2.2. 요청 변수 처리 
'##########################################################################################################

'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################
	Redim ImportData(1)
	ImportData(0) = Trim(Request("txtApNo"))
	importData(1) = "NT"

	Set iPAPG005 = Server.CreateObject("PAPG005.cALkUpOpenApSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	    
	call iPAPG005.A_LOOKUP_OPEN_AP_SVR(gStrGlobalCollection, ImportData ,ExportData1,ExportData2)

	If CheckSYSTEMError(Err,True) = True Then
		Set iPAPG005 = Nothing		
		Response.End 
	End If
	    
	Set iPAPG005 = Nothing 

	lgCurrency = ConvSPChars(ExportData1(doccur))

'#########################################################################################################
'												2.4. HTML 결과 생성부 
'##########################################################################################################
	Response.Write "<Script Language=vbscript>   " & vbcr
	Response.Write " With parent.frm1            " & vbcr														'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "	.txtDealBpCd.Value      =   """ & ConvSPChars(ExportData1(BpCd)) & """			" & vbcr
	Response.Write "	.txtDealBpNm.Value      =   """ & ConvSPChars(ExportData1(BpNm)) & """			" & vbcr
	Response.Write "	.txtDeptCd.Value        =   """ & ConvSPChars(ExportData1(DeptCd)) & """		" & vbcr
	Response.Write "	.txtDeptNm.Value        =   """ & ConvSPChars(ExportData1(DeptNm)) & """		" & vbcr
	Response.Write "	.txtPayBpCd.Value       =   """ & ConvSPChars(ExportData1(PayBpCd)) & """		" & vbcr
	Response.Write "	.txtPayBpNm.Value		=   """ & ConvSPChars(ExportData1(PayBpNm)) & """		" & vbcr
	Response.Write "	.txtReportBpCd.Value    =   """ & ConvSPChars(ExportData1(ReportBpCd)) & """	" & vbcr
	Response.Write "	.txtReportBpNm.Value	=   """ & ConvSPChars(ExportData1(ReportBpNm)) & """	" & vbcr	
	Response.Write "	.txtApDt.text           =   """ & UNIDateClientFormat(ExportData1(ApDt)) & """	" & vbcr
	Response.Write "	.txtInvNo.Value         =   """ & ConvSPChars(ExportData1(InvDocNo)) & """		" & vbcr
	Response.Write "	.txtDueDt.text          =   """ & UNIDateClientFormat(ExportData1(DueDt)) & """ " & vbcr	
	Response.Write "	.txtInvDt.text          =   """ & UNIDateClientFormat(ExportData1(InvDt)) & """ " & vbcr
	Response.Write "	.txtAcctCd.Value        =   """ & ConvSPChars(ExportData1(AcctCd)) & """		" & vbcr
	Response.Write "	.txtAcctNm.Value        =   """ & ConvSPChars(ExportData1(AcctNm)) & """		" & vbcr
	Response.Write "	.txTblNo.Value          =   """ & ConvSPChars(ExportData1(BlDocNo)) & """		" & vbcr	
	Response.Write "	.txTblDt.text           =   """ & UNIDateClientFormat(ExportData1(BlDt)) & """	" & vbcr
	Response.Write "	.txtPaymTerms.Value     =   """ & ConvSPChars(ExportData1(PaymTerms)) & """		" & vbcr
	Response.Write "	.txTlcNo.Value          =   """ & ConvSPChars(ExportData1(LcDocNo)) & """		" & vbcr
'	Response.Write "	.cboApType.Value        =   """ & ConvSPChars(ExportData1(ApType)) & """		" & vbcr	
	Response.Write "	.txtLcDt.text           =   """ & UNIDateClientFormat(ExportData1(ApLcDt)) & """" & vbcr
	Response.Write "	.txtPayDur.Value		=	""" & UniConvNum(ExportData1(PayDur), 0) & """		" & vbcr		
	Response.Write "	.txtPayMethCd.Value		=   """ & ConvSPChars(ExportData1(PayMeth)) & """		" & vbcr			
	Response.Write "	.txtPayMethNm.Value		=   """ & ConvSPChars(ExportData1(PayMethNm)) & """		" & vbcr			
	Response.Write "	.txtPayTypeCd.Value     =   """ & ConvSPChars(ExportData1(PaymType)) & """		" & vbcr	
	Response.Write "	.txtPayTypeNm.Value     =   """ & ConvSPChars(ExportData1(PaymTypeNm)) & """	" & vbcr		
	Response.Write "	.txtDocCur.Value        =   """ & ConvSPChars(ExportData1(DocCur)) & """		" & vbcr

	Response.Write "    .txtNetAmt.Text         =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(NetAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbcr
	Response.Write "    .txtBalAmt.Text         =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(BalAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbcr

	If gIsShowLocal <> "N" Then		
		Response.Write "	.txtXchRate.Text    =   """ & UNINumClientFormat(ExportData1(XchRate), ggExchRate.DecPoint, 0) & """		" & vbcr
		Response.Write "	.txtNetLocAmt.Text  =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(NetLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """	" & vbcr
		Response.Write "	.txtBalLocAmt.Text  =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(BalLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") & """" & vbcr
	Else
		Response.Write "	.txtXchRate.Value   =   """ & UNINumClientFormat(ExportData1(XchRate), ggExchRate.DecPoint, 0) & """		" & vbcr
		Response.Write "	.txtNetLocAmt.Value =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(NetLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """	" & vbcr
		Response.Write "	.txtBalLocAmt.Value =   """ & UNIConvNumDBToCompanyByCurrency(ExportData1(BalLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") & """" & vbcr
	End IF
		
	Response.Write "	.txtGlNo.Value          =   """ & ConvSPChars(ExportData1(GlNo)) & """			" & vbcr
	Response.Write "	.txtGldt.Text           =   """ & UNIDateClientFormat(ExportData1(Gldt)) & """	" & vbcr									
	Response.Write "	.txtDesc.Value	        =   """ & ConvSPChars(ExportData1(ApDesc)) & """		" & vbcr				
	Response.Write "	.htxtApNo.value			=   """ & Request("txtApNo") & """						" & vbcr

	Response.Write " End With					 " & vbcr		    
	Response.Write " Parent.DbQueryOk			 " & vbcr
	Response.write "</Script>				     " & vbcr  
%>    


