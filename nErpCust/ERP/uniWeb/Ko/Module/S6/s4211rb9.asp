<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4211rb9.asp																*
'*  4. Program Name         : 통관상세정보(통관현황조회에서)											*
'*  5. Program Desc         : 통관상세정보(통관현황조회에서)											*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : KIm Hyungsuk																*
'* 10. Modifier (Last)      : Park insik																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Dim lgADF                                                                  
Dim lgstrRetMsg                                                            
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   
                                                         
Dim lgSelectList
Dim lgSelectFrom
Dim arrRsVal(80)								

On Error Resume Next
Err.Clear
    
Call LoadBasisGlobalInf()

Call HideStatusWnd
lgSelectList= ""
	
lgSelectList = lgSelectList & " A.CC_NO	, A.SO_NO,	A.IV_NO,A.IV_DT, A.ED_NO,A.ED_DT,A.EP_NO,A.EP_DT,A.VESSEL_NM,	"
lgSelectList = lgSelectList	& "	A.SHIP_FIN_DT,A.WEIGHT_UNIT,A.LC_DOC_NO,A.LC_AMEND_SEQ,	"
lgSelectList = lgSelectList & " A.GROSS_WEIGHT,A.NET_WEIGHT, A.CUR,A.TOT_PACKING_CNT, "
lgSelectList = lgSelectList & " A.XCH_RATE,A.USD_XCH_RATE,A.CUR,A.DOC_AMT,A.LOC_AMT, "
lgSelectList = lgSelectList & " A.CUR,A.FOB_DOC_AMT,A.FOB_LOC_AMT,A.INCOTERMS,G.MINOR_NM, "
lgSelectList = lgSelectList & " A.SALES_GRP,D.SALES_GRP_NM,A.PAY_METH,H.MINOR_NM,A.PAY_DUR, "
lgSelectList = lgSelectList & " A.APPLICANT,B.BP_NM,A.BENEFICIARY,I.BP_NM,A.AGENT,J.BP_NM, "
lgSelectList = lgSelectList & " A.MANUFACTURER,K.BP_NM,A.LOADING_PORT,F.MINOR_NM,A.LOADING_CNTRY,L.COUNTRY_NM, "
lgSelectList = lgSelectList & " A.DISCHGE_PORT,E.MINOR_NM,A.DISCHGE_CNTRY,M.COUNTRY_NM,A.ORIGIN,N.MINOR_NM,	"
lgSelectList = lgSelectList & " A.ORIGIN_CNTRY,O.COUNTRY_NM,A.FINAL_DEST,A.REPORTER,P.BP_NM, "
lgSelectList = lgSelectList & " A.RETURN_APPL,Q.BP_NM,A.RETURN_OFFICE,R.MINOR_NM,A.ED_TYPE,C.MINOR_NM, "
lgSelectList = lgSelectList & " A.CUSTOMS,S.MINOR_NM,A.TRANS_FORM,T.MINOR_NM,A.PACKING_TYPE,U.MINOR_NM, "
lgSelectList = lgSelectList & " A.TRANS_REP_CD,V.BP_NM,A.TRANS_METHOD,W.MINOR_NM,A.TRANS_FROM_DT,A.TRANS_TO_DT, "
lgSelectList = lgSelectList & " A.INSP_CERT_NO,A.INSP_CERT_DT,A.QUAR_CERT_NO,A.QUAR_CERT_DT,A.DEVICE_PLCE, "
lgSelectList = lgSelectList & " A.REMARK1,A.REMARK2,A.REMARK3 "
lgSelectFrom =	" LEFT OUTER JOIN B_MINOR U ON ( A.PACKING_TYPE = U.MINOR_CD  AND  U. MAJOR_CD = " & FilterVar("B9007", "''", "S") & ") " 
lgSelectFrom =lgSelectfrom&" LEFT OUTER JOIN B_BIZ_PARTNER V ON ( A.TRANS_REP_CD = V.BP_CD) " 
lgSelectFrom =lgSelectfrom&" LEFT OUTER JOIN B_MINOR W ON ( A.TRANS_METHOD = W.MINOR_CD  AND  W.MAJOR_CD = " & FilterVar("S9011", "''", "S") & ") " 

Call FixUNISQLData()
Call QueryData()

Sub FixUNISQLData()	
	
	Dim strMode		
	Dim strVal		
		
	Redim UNISqlId(1)                   
    Redim UNIValue(0,2)					

    UNISqlId(0) = "S4211RA901"  
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList         '☜: Select list
	UNIValue(0,1) = lgSelectFrom
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strVal = ""
	strMode = Request("txtMode")									'☜ : 현재 상태를 받음	
	if strMode =CStr(UID_M0001) then								'☜: 현재 조회/Prev/Next 요청을 받음 
		Err.Clear													
		If Request("txtCCNo") = "" Then								'⊙: 조회를 위한 값이 들어왔는지 체크 
			Call ServerMesgBox("조회 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
			Response.End 
		End If
	end if
	
	If Len(Trim(Request("txtCCNo"))) Then
    	strVal = " " & FilterVar(Request("txtCCNo"), "''", "S") & " "
    	
    Else
    	strVal = "''"
    End If

	UNIValue(0,2) = strVal   

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode 
	
End Sub


Sub QueryData()
	
	Dim iStr
	Dim iCnt
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs0) '* : Record Set 의 갯수 조정 
    
	iStr = Split(lgstrRetMsg,gColSep)
    
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
    
	Else    
	
		rs0.MoveFirst
		iCnt =0
	
		For iCnt=0 to 80
		
			arrRsVal(iCnt) =rs0(iCnt)
						
		Next

		rs0.Close
		Set rs0 = Nothing 
		Set lgADF = Nothing        
	
	End If
    
End Sub

%>

<Script Language=vbscript>   	
   	
   	With parent.frm1
		Dim strDt
		'Tab 1 : Local L/C 일반정보 
		.txtCCNo1.value			= "<%=ConvSPChars(arrRsVal(0))%>"	
		.txtSONo.value			= "<%=ConvSPChars(arrRsVal(1))%>"
		.txtIVNo.value			= "<%=ConvSPChars(arrRsVal(2))%>" 
		.txtIVDt.text			= "<%=UNIDateClientFormat(arrRsVal(3))%>" 
		.txtEDNo.value			= "<%=ConvSPChars(arrRsVal(4))%>"
		.txtEDDt.text			= "<%=UNIDateClientFormat(arrRsVal(5))%>"
		.txtEPNo.value			= "<%=ConvSPChars(arrRsVal(6))%>"
		.txtEPDt.text			= "<%=UNIDateClientFormat(arrRsVal(7))%>"
		.txtVesselNm.value		= "<%=ConvSPChars(arrRsVal(8))%>"
		.txtShipFinDt.text		= "<%=UNIDateClientFormat(arrRsVal(9))%>"
		.txtWeightUnit.value	= "<%=ConvSPChars(arrRsVal(10))%>"
		.txtLCDocNo.value		= "<%=ConvSPChars(arrRsVal(11))%>"
		.txtLCAmendSeq.value	= "<%=ConvSPChars(arrRsVal(12))%>"
		.txtGrossWeight.text	= "<%=UNINumClientFormat(arrRsVal(13), ggQty.DecPoint, 0)%>"
		.txtNetWeight.text		= "<%=UNINumClientFormat(arrRsVal(14), ggQty.DecPoint, 0)%>"
		.txtCurrency.value		= "<%=ConvSPChars(arrRsVal(15))%>"
		.txtTotPackingCnt.text	= "<%=UNINumClientFormat(arrRsVal(16), ggQty.DecPoint, 0)%>"
		.txtXchRate.text		= "<%=UNINumClientFormat(arrRsVal(17), ggExchRate.DecPoint, 0)%>"
		.txtUSDXchRate.text			= "<%=UNINumClientFormat(arrRsVal(18), ggExchRate.DecPoint, 0)%>"
		.txtCCCurrency.value	= "<%=ConvSPChars(arrRsVal(19))%>"
		.txtDocAmt.text			= "<%=UNINumClientFormat(arrRsVal(20), ggAmtOfMoney.DecPoint, 0)%>"
		.txtLocAmt.text			= "<%=UNINumClientFormat(arrRsVal(21), ggAmtOfMoney.DecPoint, 0)%>"
		.txtFobCurrency.value	= "<%=ConvSPChars(arrRsVal(22))%>"
		.txtFOBDocAmt.text		= "<%=UNINumClientFormat(arrRsVal(23), ggAmtOfMoney.DecPoint, 0)%>"
		.txtFOBLocAmt.text		= "<%=UNINumClientFormat(arrRsVal(24), ggAmtOfMoney.DecPoint, 0)%>"
		.txtIncoTerms.value			= "<%=ConvSPChars(arrRsVal(25))%>"
		.txtIncoTermsNm.value		= "<%=ConvSPChars(arrRsVal(26))%>"
		.txtSalesGroup.value		= "<%=ConvSPChars(arrRsVal(27))%>"
		.txtSalesGroupNm.value		= "<%=ConvSPChars(arrRsVal(28))%>"
		.txtPayTerms.value			= "<%=ConvSPChars(arrRsVal(29))%>"
		.txtPayTermsNm.value		= "<%=ConvSPChars(arrRsVal(30))%>"
		.txtPayDur.text				= "<%=ConvSPChars(arrRsVal(31))%>"
		.txtApplicant.value		= "<%=ConvSPChars(arrRsVal(32))%>"
		.txtApplicantNm.value	= "<%=ConvSPChars(arrRsVal(33))%>"
		.txtBeneficiary.value	= "<%=ConvSPChars(arrRsVal(34))%>"
		.txtBeneficiaryNm.value = "<%=ConvSPChars(arrRsVal(35))%>"
		.txtAgent.value			= "<%=ConvSPChars(arrRsVal(36))%>"
		.txtAgentNm.value		= "<%=ConvSPChars(arrRsVal(37))%>"
		.txtManufacturer.value	= "<%=ConvSPChars(arrRsVal(38))%>"
		.txtManufacturerNm.value = "<%=ConvSPChars(arrRsVal(39))%>"
		
		'Tab 2
		.txtLoadingPort.value		= "<%=ConvSPChars(arrRsVal(40))%>"
		.txtLoadingPortNm.value		= "<%=ConvSPChars(arrRsVal(41))%>"
		.txtLoadingCntry.value		= "<%=ConvSPChars(arrRsVal(42))%>"
		.txtLoadingCntryNm.value	= "<%=ConvSPChars(arrRsVal(43))%>"
		.txtDischgePort.value		= "<%=ConvSPChars(arrRsVal(44))%>"
		.txtDischgePortNm.value		= "<%=ConvSPChars(arrRsVal(45))%>"
		.txtDischgeCntry.value		= "<%=ConvSPChars(arrRsVal(46))%>"
		.txtDischgeCntryNm.value	= "<%=ConvSPChars(arrRsVal(47))%>"
		.txtOrigin.value			= "<%=ConvSPChars(arrRsVal(48))%>"
		.txtOriginNm.value			= "<%=ConvSPChars(arrRsVal(49))%>"
		.txtOriginCntry.value		= "<%=ConvSPChars(arrRsVal(50))%>"
		.txtOriginCntryNm.value		= "<%=ConvSPChars(arrRsVal(51))%>"
		.txtFinalDest.value			= "<%=ConvSPChars(arrRsVal(52))%>"
		.txtReporter.value			= "<%=ConvSPChars(arrRsVal(53))%>"
		.txtReporterNm.value		= "<%=ConvSPChars(arrRsVal(54))%>"
		.txtReturnAppl.value		= "<%=ConvSPChars(arrRsVal(55))%>"
		.txtReturnApplNm.value		= "<%=ConvSPChars(arrRsVal(56))%>"
		.txtReturnOffice.value		= "<%=ConvSPChars(arrRsVal(57))%>"
		.txtReturnOfficeNm.value	= "<%=ConvSPChars(arrRsVal(58))%>"
		.txtEDType.value		= "<%=ConvSPChars(arrRsVal(59))%>"
		.txtEDTypeNm.value		= "<%=ConvSPChars(arrRsVal(60))%>"
		.txtCustoms.value			= "<%=ConvSPChars(arrRsVal(61))%>"
		.txtCustomsNm.value			= "<%=ConvSPChars(arrRsVal(62))%>"
		.txtTransForm.value			= "<%=ConvSPChars(arrRsVal(63))%>"
		.txtTransFormNm.value		= "<%=ConvSPChars(arrRsVal(64))%>"
		.txtPackingType.value		= "<%=ConvSPChars(arrRsVal(65))%>"
		.txtPackingTypeNm.value		= "<%=ConvSPChars(arrRsVal(66))%>"
		.txtTransRepCd.value		= "<%=ConvSPChars(arrRsVal(67))%>"
		.txtTransRepNm.value		= "<%=ConvSPChars(arrRsVal(68))%>"
		.txtTransMeth.value			= "<%=ConvSPChars(arrRsVal(69))%>"
		.txtTransMethNm.value		= "<%=ConvSPChars(arrRsVal(70))%>"		
		.txtTransFromDt.text		= "<%=UNIDateClientFormat(arrRsVal(71))%>"
		.txtTransToDt.text			= "<%=UNIDateClientFormat(arrRsVal(72))%>"
		.txtInspCertNo.value		= "<%=ConvSPChars(arrRsVal(73))%>"		
		.txtInspCertDt.text			= "<%=UNIDateClientFormat(arrRsVal(74))%>"
		.txtQuarCertNo.value		= "<%=ConvSPChars(arrRsVal(75))%>"		
		.txtQuarCertDt.text			= "<%=UNIDateClientFormat(arrRsVal(76))%>"	
		.txtDevicePlce.value		= "<%=ConvSPChars(arrRsVal(77))%>"
		.txtRemark1.value			= "<%=ConvSPChars(arrRsVal(78))%>"
		.txtRemark2.value			= "<%=ConvSPChars(arrRsVal(79))%>"
		.txtRemark3.value			= "<%=ConvSPChars(arrRsVal(80))%>"

		Call parent.DbQueryOk()														'☜: 조회가 성공 
		.txtHCCNo.value = "<%=ConvSPChars(Request("txtCCNo"))%>"
	End With

</Script>	

<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
		
