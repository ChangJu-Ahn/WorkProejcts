<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4713mb5.asp
'*  4. Program Name         : Lookup Production Info.
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001/12/12
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf

On Error Resume Next

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0										'DBAgent Parameter 선언 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

Dim strItemCd
Dim strOprNo
Dim strRoutNo
Dim strFlag

Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "p4713mb5"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtOprNo")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		'Call DisplayMsgBox("189300", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		%>
		<Script Language=vbscript>
			Call parent.LookUpOrderDetailFail(CInt("<%=Request("txtRow")%>"))
		</Script>	
		<%		
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>
<Script Language=vbscript>

	With parent.frm1.vspdData1

		.Row = CLng("<%=Request("txtRow")%>")


			.Col = parent.C_JobCd
			.text = "<%=ConvSPChars(rs0("job_cd"))%>"
			.Col = parent.C_JobNm
			.text = "<%=ConvSPChars(rs0("job_cd"))%>"
			.Col = parent.C_WcCd
			.text = "<%=ConvSPChars(rs0("wc_cd"))%>"
			.Col = parent.C_WcNm
			.text = "<%=ConvSPChars(rs0("wc_nm"))%>"
			.Col = parent.C_OrderStatus
			.text = "<%=ConvSPChars(rs0("order_status"))%>"
			.Col = parent.C_OrderStatusDesc
			.text = "<%=ConvSPChars(rs0("order_status"))%>"
			.Col = parent.C_PlanStartDt
			.text = "<%=UNIDateClientFormat(rs0("plan_start_dt"))%>"
			.Col = parent.C_PlanComptDt
			.text = "<%=UNIDateClientFormat(rs0("plan_compt_dt"))%>"
			.Col = parent.C_ReleaseDt
			.text = "<%=UNIDateClientFormat(rs0("release_dt"))%>"
			.Col = parent.C_RealStartDt
			.text = "<%=UNIDateClientFormat(rs0("real_start_dt"))%>"
			
			Call parent.LookUpOrderDetailSuccess(CInt("<%=Request("txtRow")%>"))

	End With
	
<%
	rs0.Close
	Set rs0 = Nothing
%>
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
