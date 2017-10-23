<%@ LANGUAGE=VBSCript %>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1410mb2.asp 
'*  4. Program Name         : ECN Management
'*  5. Program Desc         : Manage ECN Master
'*  6. Modified date(First) : 2003/03/06
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf

Dim oPP1S411

Dim strMode													'☆ : Lookup 용 코드 저장 변수 
Dim lgIntFlgMode
Dim iCommandSent

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter 선언 
Dim rs0														'DBAgent Parameter 선언 

Dim strECNNo
Dim strECNDesc
Dim strReasonCd
Dim strStatus
Dim strValidFromDt
Dim strValidToDt
Dim strEBomFlg
Dim strEBomDt
Dim strMBomFlg
Dim strMBomDt
Dim strRemark

Dim I1_P_ECN_Master
Dim E1_P_ECN_No

ReDim I1_P_ECN_Master(10)
Const C_I1_ECN_No		= 0		'Update시 필요 
Const C_I1_ECN_Desc		= 1
Const C_I1_Reason_Cd	= 2
Const C_I1_Valid_From_Dt= 3
Const C_I1_Valid_To_Dt	= 4
Const C_I1_Status		= 5
Const C_I1_EBom_Flg		= 6
Const C_I1_EBom_Dt		= 7
Const C_I1_MBom_Flg		= 8
Const C_I1_MBom_Dt		= 9
Const C_I1_Remark		= 10

Call HideStatusWnd

On Error Resume Next
Err.Clear  

	strMode = Request("txtMode")							'☜ : 현재 상태를 받음 
	lgIntFlgMode = CInt(Request("txtFlgMode"))				'☜: 저장시 Create/Update 판별 
	
	strECNNo		= UCase(Trim(Request("txtECNNo1")))
	strECNDesc		= Request("txtECNDesc1")
	strReasonCd		= UCase(Trim(Request("txtReasonCd")))
	strValidFromDt	= Request("txtValidFromDt")
	strValidToDt	= Request("txtValidToDt")
	strStatus		= Request("cboStatus")
	strEBomFlg		= "N"
	strEBomDt		= ""
	strMBomFlg		= "N"
	strMBomDt		= ""
	If Request("txtEBomFlg") <> "" Then
		strEBomFlg		= Request("txtEBomFlg")
		strEBomDt		= Request("txtEBomDt")
	End if

	If Request("txtMBomFlg") <> "" Then
		strMBomFlg		= Request("txtMBomFlg")
		strMBomDt		= Request("txtMBomDt")
	End if
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If
	strRemark		= Request("txtRemark")
	
'변경근거코드 체크 - CREATE일때만?

	'--------------------------------------------
	' 변경근거가 존재하는지 체크 
	'--------------------------------------------
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "s0000qa000"
	
	UNIValue(0, 0) = FilterVar("P1402","''","S")		'major_cd
	UNIValue(0, 1) = FilterVar(strReasonCd,"''","S")	'minor_cd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("182803", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=vbscript>			" & vbCr
		Response.Write "	parent.frm1.txtReasonCd.focus()	" & vbCr															
		Response.Write "</Script>							" & vbCr
		Response.End
	End If

	I1_P_ECN_Master(C_I1_ECN_No)		= strECNNo
	I1_P_ECN_Master(C_I1_ECN_Desc)		= strECNDesc
	I1_P_ECN_Master(C_I1_Reason_Cd)		= strReasonCd
	I1_P_ECN_Master(C_I1_Valid_From_Dt) = strValidFromDt
	I1_P_ECN_Master(C_I1_Valid_To_Dt)	= strValidToDt
	I1_P_ECN_Master(C_I1_Status)		= strStatus
	I1_P_ECN_Master(C_I1_EBom_Flg)		= strEBomFlg
	I1_P_ECN_Master(C_I1_EBom_Dt)		= strEBomDt
	I1_P_ECN_Master(C_I1_MBom_Flg)		= strMBomFlg
	I1_P_ECN_Master(C_I1_MBom_Dt)		= strMBomDt
	I1_P_ECN_Master(C_I1_Remark)		= strRemark

	Set oPP1S411 = Server.CreateObject("PP1S411.cPMngEcn")

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

    Call oPP1S411.P_MANAGE_ECN(gStrGlobalCollection, _
							iCommandSent, _
							I1_P_ECN_Master, _
							E1_P_ECN_No)

    If CheckSYSTEMError(Err,True) = True Then
       Set oPP1S411 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.End		
    End If
	
	Set oPP1S411 = Nothing
	
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr	
	Response.Write "		.frm1.txtECNNo.value = """ & ConvSPChars(E1_P_ECN_No) & """" & vbCr 'sjdklfjsd
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr

	Response.End																				'☜: Process End

	'==============================================================================
	' 사용자 정의 서버 함수 
	'==============================================================================

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>	" & vbCr
	Response.Write "										" & vbCr
	Response.Write "</SCRIPT>								" & vbCr

%>