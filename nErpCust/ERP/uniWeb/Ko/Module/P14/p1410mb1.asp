<%@ LANGUAGE=VBSCript %>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1410mb1.asp 
'*  4. Program Name         : ECN Management
'*  5. Program Desc         : Lookup ECN Master
'*  6. Modified date(First) : 2003/03/05
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim oPP1S412

Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 

Dim strPrevNextFlg			'String
Dim strECNNo				'String
Dim I1_P_ECN_No				'String
Dim E1_P_ECN_Master
Dim strStatusCodeOfPrevNext

Const C_E1_ECN_No		= 0
Const C_E1_ECN_Desc		= 1
Const C_E1_Reason_Cd	= 2
Const C_E1_Issued_By	= 3
Const C_E1_Valid_From_Dt= 4
Const C_E1_Valid_To_Dt	= 5
Const C_E1_Status		= 6
Const C_E1_EBom_Flg		= 7
Const C_E1_EBom_Dt		= 8
Const C_E1_MBom_Flg		= 9
Const C_E1_MBom_Dt		= 10
Const C_E1_Remark		= 11

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
	strPrevNextFlg = Request("PrevNextFlg")
	strECNNo = UCase(Trim(Request("txtECNNo")))
 
	%>
	<Script Language=vbscript>
		parent.frm1.txtECNDesc.value = ""
	</Script>
	<%

	I1_P_ECN_No = strECNNo

	Set oPP1S412 = Server.CreateObject("PP1S412.cPLkupEcnSvr")

    If CheckSYSTEMError(Err,True) = True Then
		Response.End 
    End If

	Call oPP1S412.P_LOOK_UP_ECN_SVR(gStrGlobalCollection, _
								strPrevNextFlg, _
								I1_P_ECN_No, _
								E1_P_ECN_Master, _
								strStatusCodeOfPrevNext)

    If CheckSYSTEMError(Err,True) = True Then
       Set oPP1S412 = Nothing
       Response.End 
    End If
    
    Set oPP1S412 = Nothing															'☜: Unload Comproxy
    
	If (strStatusCodeOfPrevNext = "900011" Or strStatusCodeOfPrevNext = "900012") Then
		Call DisplayMsgBox(strStatusCodeOfPrevNext, vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	End If

	Response.Write "<Script Language = VBScript> " & vbCr
	Response.Write "Dim LngRow " & vbCr
	
	Response.Write "With parent.frm1 " & vbCr
	Response.Write "	.txtECNNo.value			= """ & ConvSPChars(UCase(Trim(E1_P_ECN_Master(C_E1_ECN_No)))) & """" & vbCr
	Response.Write "	.txtECNDesc.value		= """ & ConvSPChars(E1_P_ECN_Master(C_E1_ECN_Desc)) & """" & vbCr
	Response.Write "	.txtECNNo1.value		= """ & ConvSPChars(UCase(Trim(E1_P_ECN_Master(C_E1_ECN_No)))) & """" & vbCr
	Response.Write "	.txtECNDesc1.value		= """ & ConvSPChars(E1_P_ECN_Master(C_E1_ECN_Desc)) & """" & vbCr
	Response.Write "	.txtReasonCd.value		= """ & ConvSPChars(UCase(Trim(E1_P_ECN_Master(C_E1_Reason_Cd)))) & """" & vbCr
	Response.Write "	.txtIssuedBy.value		= """ & ConvSPChars(E1_P_ECN_Master(C_E1_Issued_By)) & """" & vbCr
	Response.Write "	.cboStatus.value		= """ & ConvSPChars(E1_P_ECN_Master(C_E1_Status)) & """" & vbCr
	Response.Write "	.txtEBomFlg.value		= """ & ConvSPChars(UCase(Trim(E1_P_ECN_Master(C_E1_EBom_Flg)))) & """" & vbCr
	Response.Write "	.txtEBomDt.text			= """ & UNIDateClientFormat(E1_P_ECN_Master(C_E1_EBom_Dt)) & """" & vbCr
	Response.Write "	.txtMBomFlg.value		= """ & ConvSPChars(UCase(Trim(E1_P_ECN_Master(C_E1_MBom_Flg)))) & """" & vbCr
	Response.Write "	.txtMBomDt.text			= """ & UNIDateClientFormat(E1_P_ECN_Master(C_E1_MBom_Dt)) & """" & vbCr
	Response.Write "	.txtValidFromDt.text	= """ & UNIDateClientFormat(E1_P_ECN_Master(C_E1_Valid_From_Dt)) & """" & vbCr
	Response.Write "	.txtValidToDt.text		= """ & UNIDateClientFormat(E1_P_ECN_Master(C_E1_Valid_To_Dt)) & """" & vbCr
	Response.Write "	.txtRemark.value		= """ & ConvSPChars(E1_P_ECN_Master(C_E1_Remark)) & """" & vbCr
	
	Response.Write "	parent.lgNextNo = """"" & vbCr		' 다음 키 값 넘겨줌 
	Response.Write "	parent.lgPrevNo = """"" & vbCr		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 

	If UCase(Trim(E1_P_ECN_Master(C_E1_EBom_Flg))) = "Y" OR UCase(Trim(E1_P_ECN_Master(C_E1_MBom_Flg))) = "Y" Then
		Response.Write "	parent.blnBomFlg	= True " & vbCr
	Else
		Response.Write "	parent.blnBomFlg	= False " & vbCr
	End If

	Response.Write "	parent.DbQueryOk " & vbCr								'☜: 조회가 성공 
	Response.Write "End With " & vbCr
	Response.Write "</Script> " & vbCr    

	Response.End																	'☜: Process End

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER> " & vbCr
	Response.Write "" & vbCr
	Response.Write "</SCRIPT> " & vbCr
%>

