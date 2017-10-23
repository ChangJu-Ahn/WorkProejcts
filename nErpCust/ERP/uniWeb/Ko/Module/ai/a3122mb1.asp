<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a3104mb1
'*  4. Program Name         : 가수금내역조회 
'*  5. Program Desc         : 가수금내역조회 
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/10/13
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : 김희정 
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 

'#########################################################################################################
'												1. Include
'##########################################################################################################
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'#########################################################################################################
'												2. 조건부 
'##########################################################################################################
																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
On Error Resume Next														'☜: 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

'#########################################################################################################
'												2.1 조건 체크 
'##########################################################################################################
If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then										'☜: 조회 전용 Biz 이므로 다른값은 그냥 종료함 
	Call ServerMesgBox("700118", vbInformation, I_MKSCRIPT)					'⊙: 조회 전용인데 다른 상태로 요청이 왔을 경우, 필요없으면 빼도 됨, 메세지는 ID값으로 사용해야 함 
	Response.End 
ElseIf Request("txtRcptNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call ServerMesgBox("700112", vbInformation, I_MKSCRIPT)					'⊙:
	Response.End 
End If

'#########################################################################################################
'												2. 업무 처리 수행부 
'##########################################################################################################

'#########################################################################################################
'												2.1. 변수, 상수 선언 
'##########################################################################################################
Dim iArrData
Dim iGData
Dim lgStrPrevKey
Dim iLngRow
Dim LngMaxRow
Dim iARcptItemSeq
Dim iPARG020
Dim iStrData
Dim lgCurrency
Dim iRcptNo
Dim iRcptInputType

Const RcptNo = 0
Const JnlCd = 1
Const JnlNm = 2
Const ConfFg = 3
Const DeptCd = 4
Const DeptNm = 5
Const RcptDt = 6
Const BpCd = 7
Const BpNm = 8
Const RefNo = 9
Const DocCur = 10
Const XchRate = 11
Const RcptAmt = 12
Const RcptLocAmt = 13
Const BnkChgAmt = 14
Const BnkChgLocAmt = 15
Const AllcAmt = 16
Const AllclocAmt = 17
Const Adjustamt = 18
Const AdjustLocAmt = 19
Const BalAmt = 20
Const BalLocAmt = 21
Const TempGlNo = 22
Const GlNo = 23
Const RcptDesc = 24
Const Project = 25
Const GLDt = 26
'#########################################################################################################
'												2.2. 요청 변수 처리 
'##########################################################################################################
	lgStrPrevKey = Request("lgStrPrevKey")
'#########################################################################################################
'												2.3. 업무 처리 
'##########################################################################################################

	If lgStrPrevKey = "" Then
		iARcptItemSeq = 0
	Else
		iARcptItemSeq = lgStrPrevKey
	End If
	Set iPARG020 = Server.CreateObject("PARG020.cALkUpRcSvr")

	If CheckSYSTEMError(Err, True) = True Then					
		Response.End 
	End If    
		
	iRcptNo = Trim(Request("txtRcptNo"))
	iRcptInputType = "RT"

	Call iPARG020.LOOKUP_RCPT_SVR(gStrGloBalCollection, iARcptItemSeq, iRcptNo,iRcptInputType, iArrData, iGData)
	
	If CheckSYSTEMError(Err, True) = True Then					
	   Set iPARG020 = Nothing
	   Response.End 
	End If    

	lgCurrency = iArrDAta(DocCur)
	Response.Write ConvSPChars(iArrDAta(ConfFg)) & "&&" & ConvSPChars(iArrData(TempGlNo))	 & "&& " & ConvSPChars(iArrData(GlNo))
	
	Response.Write "<Script Language=vbscript>  " & vbcr
	Response.Write " With parent.frm1           " & vbcr														'☜: 화면 처리 ASP 를 지칭함 
	Response.Write ".txtRcptNo.Value		= """ & ConvSPChars(iArrData(RcptNo))			& """ " & vbcr
	Response.Write ".txtRcptType.value		= """ & ConvSPChars(iArrData(JnlCd))			& """ " & vbcr
	Response.Write ".txtRcptTypeNm.value	= """ & ConvSPChars(iArrData(JnlNm))			& """ " & vbcr
	Response.Write ".txtDeptNm.Value		= """ & ConvSPChars(iArrData(DeptNm))			& """ " & vbcr
	Response.Write ".txtDept.Value			= """ & ConvSPChars(iArrData(DeptCd))			& """ " & vbcr
	Response.Write ".fpDateTime1.Text       = """ & UNIDateClientFormat(iArrData(RcptDt))	& """ " & vbcr
	Response.Write ".txtBpCd.Value			= """ & ConvSPChars(iArrData(BpCd))				& """ " & vbcr
	Response.Write ".txtBpNm.Value			= """ & ConvSPChars(iArrData(BpNm))				& """ " & vbcr
	Response.Write ".txtRefNo.value			= """ & ConvSPChars(iArrDAta(RefNo))			& """ " & vbcr
	Response.Write ".txtDocCur.Value		= """ & ConvSPChars(iArrDAta(DocCur))			& """ " & vbcr
	Response.Write ".txtXchRate.Value		= """ & UNINumClientFormat(iArrDAta(XchRate), ggExchRate.DecPoint, 0)			& """ " &vbcr

	Response.Write ".txtRcptAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(RcptAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """ " &vbcr
	Response.Write ".txtRcptLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(RcptLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """ " & vbcr

	Response.Write ".txtClsAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AllcAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """ " & vbcr 
	Response.Write ".txtClsLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AllcLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """ " & vbcr

	Response.Write ".txtSttlAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AdjustAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")		& """ " & vbcr 
	Response.Write ".txtSttlLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(AdjustLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")	& """ " & vbcr

	Response.Write ".txtBalAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(iArrData(BalAmt), lgCurrency,ggAmtOfMoneyNo, "X" , "X")			& """ " & vbcr 
	Response.Write ".txtBalLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(iArrData(BalLocAmt), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")		& """ " & vbcr
	Response.Write ".htxtTempGlNO.Value			= """ & ConvSPChars(iArrData(TempGlNo))				& """ " & vbcr
	Response.Write ".txtGlNo.Value			= """ & ConvSPChars(iArrData(GlNo))				& """ " & vbcr
	Response.Write ".txtDesc.value			= """ & ConvSPChars(iArrDAta(RcptDesc))			& """ " & vbcr
	Response.Write ".txtProject.value		= """ & ConvSPChars(iArrDAta(Project))			& """ " & vbcr	
	'If ConvSPChars(iArrDAta(ConfFg)) = "C"	 Then
	'	Response.Write ".chkConfFg.checked		= True " & vbcr
	'	Response.Write ".txtConfFg.value		= ""C"" " & vbcr
	'Else
		if ConvSPChars(iArrData(GlNo))= "" and ConvSPChars(iArrData(TempGlNo))="" then		'temp_gl 
			if ConvSPChars(iArrDAta(ConfFg)) = "C"	 then
				Response.Write ".chkConfFg.checked		= True " & vbcr
				Response.Write ".txtConfFg.value		= ""C"" " & vbcr
			ELse
				Response.Write ".chkConfFg.checked		= False " & vbcr
				Response.Write ".txtConfFg.value		= ""U"" " & vbcr
			ENd if
		Else
			Response.Write ".chkConfFg.checked		= True " & vbcr
			Response.Write ".txtConfFg.value		= ""C"" " & vbcr
		End if
	'End If
	if UNIDateClientFormat(iArrData(GLDt))<> ""then
		Response.Write ".htxtgldt.Value       = """ & UNIDateClientFormat(iArrData(GLDt))	& """ " & vbcr 
	End if
	Response.Write " End With					" & vbcr		    
	Response.Write " Parent.DbQueryOk			" & vbcr
	Response.Write "</Script>  " & vbcr
%>

