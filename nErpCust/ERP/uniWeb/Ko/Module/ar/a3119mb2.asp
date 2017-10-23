<%
'======================================================================================================
'*  1. Module Name          : Finance
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6102mb2
'*  4. Program Name         : 입금청산 
'*  5. Program Desc         : 입금청산의 LookUp
'*  6. Modified date(First) : 2000/09/27
'*  7. Modified date(Last)  : 2000/12/20
'*  8. Modifier (First)     : 조익성 
'*  9. Modifier (Last)      : hersheys
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

								'☜ : ASP가 캐쉬되지 않도록 한다.
								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd

On Error Resume Next														'☜: 

Dim pFp0011																	'입력/수정용 ComProxy Dll 사용 변수 
Dim pAr0071																	'조회용 ComProxy Dll 사용 변수 
Dim strCode																	'Lookup 용 코드 저장 변수 

Dim strMode																	'현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

    Err.Clear                                                               '☜: Protect system from crashing
    
    Set pAr0071 = Server.CreateObject("Ar0071.ALookupRcptAdjustSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If Err.Number <> 0 Then
		Set pAr0071 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'에러내용, 메세지타입, 스크립트유형 
		Response.End																		'☜: Process End
	End If
    
    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	pAr0071.ImprotARcptRcptNo		= Request("txtRcptNo")
	pAr0071.ImportARcptAdjustAdjustNo = lgStrPrevKey
	'pAr0071.ImportIefSuppliedSelectChar	= Request("SelectChar")
	'pAr0071.ImportBAcctDeptOrgChangeId	= gChangeOrgID
    pAr0071.ServerLocation              = ggServerIP
    
    '------------------------------------------
    'Com action area
    '------------------------------------------
    pAr0071.ComCfg = gConnectionString
    pAr0071.Execute					

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                      '에러내용, 메세지타입, 스크립트유형 
	   Set pAr0071 = Nothing													            '☜: ComProxy UnLoad
	   Response.End																			'☜: Process End
	End If
    
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAr0071.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAr0071.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAr0071.ExportErrEabSqlCodeSqlcode, _
						    pAr0071.ExportErrEabSqlCodeSeverity, _
						    pAr0071.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAr0071.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAr0071 = Nothing
		Response.End
	End If  

	'------------------------------------------
	'Result data display area
	'------------------------------------------
	' 다음키와 이전키는 존재하지 않을 경우 Blank로 보내는 로직을 수행함.
%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	With parent.frm1
<%	
	Dim lgCurrency
	lgCurrency = ConvSPChars(pAr0071.ExportARcptDocCur)
%>	
		.txtRcptNo.value	   = "<%=ConvSPChars(pAr0071.ExportARcptRcptNo)%>"

		.txtDeptCd.value	   = "<%=ConvSPChars(pAr0071.ExportBAcctDeptDeptCd)%>"
		.txtDeptNm.value	   = "<%=ConvSPChars(pAr0071.ExportBAcctDeptDeptNm)%>"
		.txtRcptDt.text		   = "<%=UNIDateClientFormat(pAr0071.ExportARcptRcptDt)%>"
		.txtBpCd.value		   = "<%=ConvSPChars(pAr0071.ExportBBizPartnerBpCd)%>"
		.txtBpNm.value		   = "<%=ConvSPChars(pAr0071.ExportBBizPartnerBpNm)%>"
		.txtRefNo.value		   = "<%=ConvSPChars(pAr0071.ExportARcptRcptNo)%>"
		.txtDocCur.value	   = "<%=ConvSPChars(pAr0071.ExportARcptDocCur)%>"		
		
		.txtRcptAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptRcptAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtRcptLocAmt.value   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptRcptLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>" 
		.txtBalAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptBalAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtBalLocAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptBalLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>" 
		
		.txtGlNo.value		   = "<%=ConvSPChars(pAr0071.ExportAGlGlNo)%>"
		.txtRcptDesc.value   = "<%=ConvSPChars(pAr0071.ExportARcptRcptDesc)%>"

		parent.lgNextNo = ""		' 다음 키 값 넘겨줌 
		parent.lgPrevNo = ""		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 
		
		parent.DbQueryOk															'☜: 조회가 성공 
	End With
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 

	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow                                                
<%      
  	For LngRow = 1 To GroupCount
%>
        strData = strData & Chr(11) & "<%=LngRow%>"	'1  C_AdjustNo
        strData = strData & Chr(11) & "<%=UNIDateClientFormat(pAr0071.ExportARcptAdjustAdjustDt(LngRow))%>" '2 C_AdjustDt
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportAAcctAcctCd(LngRow))%>"			'3  C_AcctCd
        strData = strData & Chr(11) & ""													'4  C_AcctCdPopUp
        strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportAAcctAcctNm(LngRow))%>"  		'5  C_AcctNm 
		
        strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptAdjustAdjustAmt(LngRow), lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
		strData = strData & Chr(11) & "<%=UNIConvNumDBToCompanyByCurrency(pAr0071.ExportARcptAdjustAdjustLocAmt(LngRow), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		
		strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportARcptAdjustDocDur(LngRow))%>"	'8  C_DocCur
        strData = strData & Chr(11) & ""													'9  C_DocCurPopUp
		strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportItemFPrpaymSttlGlNo(LngRow))%>"		'10 C_GlNo
		strData = strData & Chr(11) & "<%=ConvSPChars(pAr0071.ExportAdjustAGlGlNo(LngRow))%>"		'11 C_RefNo

        strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>									
        strData = strData & Chr(11) & Chr(12)
<%      
    Next
%>    
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowData strData
		
	.lgStrPrevKey = "<%=StrNextKey%>"

	.frm1.hRcptNo.value = "<%=Request("txtRcptNo")%>"

'	.DbQueryOk
				
	End With
</Script>
<%
	 
    Set pAr0071 = Nothing															'☜: Unload Comproxy

	Response.End																	'☜: Process End

End Select
%>
