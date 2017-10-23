<%
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6102mb2
'*  4. Program Name         : 채권청산 
'*  5. Program Desc         : 채권청산의 LookUp
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
																	'입력/수정용 ComProxy Dll 사용 변수 
Dim pAr0081																	'조회용 ComProxy Dll 사용 변수 
Dim strCode																	'Lookup 용 코드 저장 변수 

Dim strMode																	'현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

    Err.Clear                                                               '☜: Protect system from crashing
    
    Set pAr0081 = Server.CreateObject("Ar0089.ALookupArAdjustSvrv")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If Err.Number <> 0 Then
		Set pAr0081 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'에러내용, 메세지타입, 스크립트유형 
		Response.End																		'☜: Process End
	End If
    
    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	pAr0081.ImportAOpenArArNo		= Trim(Request("txtArNo"))
	'pAr0081.ImportIefSuppliedSelectChar	= Request("SelectChar")
	'pAr0081.ImportBAcctDeptOrgChangeId	= gChangeOrgID
    pAr0081.ServerLocation              = ggServerIP
    
    '------------------------------------------
    'Com action area
    '------------------------------------------
    pAr0081.ComCfg = gConnectionString
    pAr0081.Execute					

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                      '에러내용, 메세지타입, 스크립트유형 
	   Set pAr0081 = Nothing													            '☜: ComProxy UnLoad
	   Response.End																			'☜: Process End
	End If
    
	'------------------------------------------
	'Com action result check area(DB,internal)
	'------------------------------------------
	If Not (pAr0081.OperationStatusMessage = MSG_OK_STR) Then
	    Select Case pAr0081.OperationStatusMessage
            Case MSG_DEADLOCK_STR
                Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
            Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAr0081.ExportErrEabSqlCodeSqlcode, _
						    pAr0081.ExportErrEabSqlCodeSeverity, _
						    pAr0081.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
            Case Else
                Call DisplayMsgBox(pAr0081.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
            End Select

		Set pAr0081 = Nothing
		Response.End 
	End If

	'------------------------------------------
	'Result data display area
	'------------------------------------------
	' 다음키와 이전키는 존재하지 않을 경우 Blank로 보내는 로직을 수행함.
%>
<Script Language=vbscript>
	With parent.frm1
<%	
	Dim lgCurrency
	lgCurrency = ConvSPChars(pAr0081.ExportAOpenArDocCur)
%>	
		.txtArNo.value			= "<%=ConvSPChars(pAr0081.ExportAOpenArArNo)%>"
		.txtDeptCd.value	   = "<%=ConvSPChars(pAr0081.ExportBAcctDeptDeptCd)%>"
		.txtDeptNm.value	   = "<%=ConvSPChars(pAr0081.ExportBAcctDeptDeptNm)%>"
		.txtArDt.text		   = "<%=UNIDateClientFormat(pAr0081.ExportAOpenArArDt)%>"
		.txtBpCd.value		   = "<%=ConvSPChars(pAr0081.ExportBBizPartnerBpCd)%>"
		.txtBpNm.value		   = "<%=ConvSPChars(pAr0081.ExportBBizPartnerBpNm)%>"
		.txtRefNo.value		   = "<%=ConvSPChars(pAr0081.ExportAOpenArRefNo)%>"
		.txtDocCur.value	   = "<%=ConvSPChars(pAr0081.ExportAOpenArDocCur)%>"		
		
		.txtArAmt.value			= "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtArLocAmt.value		= "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>" 
		.txtBalAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArAmt - opAr0081.ExportAOpenArClsAmt - ExportAOpenArAdjustAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>" 
		.txtBalLocAmt.value	   = "<%=UNIConvNumDBToCompanyByCurrency(pAr0081.ExportAOpenArArLocAmt - opAr0081.ExportAOpenArClsLocAmt - ExportAOpenArAdjustLocAmt,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>".txtGlNo.value		   = "<%=pAr0081.ExportAGlGlNo%>"
		.txtGlNo.value		   = "<%=ConvSPChars(opAr0081.ExportAGlGlNo)%>"
		.txtArDesc.value		= "<%=ConvSPChars(pAr0081.ExportAOpenArArDesc)%>"

		parent.lgNextNo = ""		' 다음 키 값 넘겨줌 
		parent.lgPrevNo = ""		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 
		
		parent.DbQueryOk															'☜: 조회가 성공 
	End With
</Script>
<%
	 
    Set pAr0081 = Nothing															'☜: Unload Comproxy

	Response.End																	'☜: Process End

End Select
%>
