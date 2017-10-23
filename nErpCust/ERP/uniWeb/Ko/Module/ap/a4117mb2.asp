<%
'======================================================================================================
'*  1. Module Name          : Finance
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6102mb2
'*  4. Program Name         : 채무청산 
'*  5. Program Desc         : 채무청산의 LookUp
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
Dim pAp0061																	'조회용 ComProxy Dll 사용 변수 
Dim strCode																	'Lookup 용 코드 저장 변수 

Dim strMode																	'현재 MyBiz.asp 의 진행상태를 나타냄 

strMode = Request("txtMode")												'현재 상태를 받음 

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

    Err.Clear                                                               '☜: Protect system from crashing
    
    Set pAp0061 = Server.CreateObject("Ap0069.ALookupRcptAdjustSvr")

    '------------------------------------------
    'Com action result check area(OS,internal)
    '------------------------------------------
    If Err.Number <> 0 Then
		Set pAp0061 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'에러내용, 메세지타입, 스크립트유형 
		Response.End																		'☜: Process End
	End If
    
    '------------------------------------------
    'Data manipulate  area(import view match)
    '------------------------------------------
	pAp0061.ImprotARcptRcptNo		= Request("txtApNo")
	'pAp0061.ImportIefSuppliedSelectChar	= Request("SelectChar")
	'pAp0061.ImportBAcctDeptOrgChangeId	= gChangeOrgID
    pAp0061.ServerLocation              = ggServerIP
    
    '------------------------------------------
    'Com action area
    '------------------------------------------
    pAp0061.ComCfg = gConnectionString
    pAp0061.Execute					

	'------------------------------------------
	'Com action result check area(OS,internal)
	'------------------------------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                      '에러내용, 메세지타입, 스크립트유형 
	   Set pAp0061 = Nothing													            '☜: ComProxy UnLoad
	   Response.End																			'☜: Process End
	End If
    
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pAp0061.OperationStatusMessage = MSG_OK_STR) Then
		Select Case pAp0061.OperationStatusMessage
			Case MSG_DEADLOCK_STR
				Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
			Case MSG_DBERROR_STR
				Call DisplayMsgBox2(pAp0061.ExportErrEabSqlCodeSqlcode, _
						    pAp0061.ExportErrEabSqlCodeSeverity, _
						    pAp0061.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
			Case Else
				Call DisplayMsgBox(pAp0061.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		End Select

		Set pAp0061 = Nothing
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
	lgCurrency = ConvSPChars(pAp0061.ExportARcptDocCur)
%>	
		.txtApNo.value	   = "<%=ConvSPChars(pAp0061.ExportARcptRcptNo)%>"

		.txtDeptCd.value	   = "<%=ConvSPChars(pAp0061.ExportBAcctDeptDeptCd)%>"
		.txtDeptNm.value	   = "<%=ConvSPChars(pAp0061.ExportBAcctDeptDeptNm)%>"
		.txtApDt.text		   = "<%=UNIDateClientFormat(pAp0061.ExportARcptRcptDt)%>"
		.txtBpCd.value		   = "<%=ConvSPChars(pAp0061.ExportBBizPartnerBpCd)%>"
		.txtBpNm.value		   = "<%=ConvSPChars(pAp0061.ExportBBizPartnerBpNm)%>"
		.txtRefNo.value		   = "<%=ConvSPChars(pAp0061.ExportARcptRcptNo)%>"
		.txtDocCur.value	   = "<%=ConvSPChars(pAp0061.ExportARcptDocCur)%>"		
		
		.txtApAmt.value	    = "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptRcptAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
		.txtApLocAmt.value  = "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptRcptLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		.txtBalAmt.value	= "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptBalAmt, lgCurrency,ggAmtOfMoneyNo, "X" , "X")%>"
		.txtBalLocAmt.value	= "<%=UNIConvNumDBToCompanyByCurrency(pAp0061.ExportARcptBalLocAmt, gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")%>"
		
		.txtGlNo.value		   = "<%=ConvSPChars(pAp0061.ExportAGlGlNo)%>"
		.txtApDesc.value   = "<%=ConvSPChars(pAp0061.ExportARcptRcptDesc)%>"

		parent.lgNextNo = ""		' 다음 키 값 넘겨줌 
		parent.lgPrevNo = ""		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 
		
		parent.DbQueryOk															'☜: 조회가 성공 
	End With
</Script>
<%
	 
    Set pAp0061 = Nothing															'☜: Unload Comproxy

	Response.End																	'☜: Process End

End Select
%>
