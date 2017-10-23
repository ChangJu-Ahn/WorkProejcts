<%@ LANGUAGE=VBSCript %>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b22mb1.asp 
'*  4. Program Name         : Called By B3B22MA1 (Class Management)
'*  5. Program Desc         : Lookup Class Information
'*  6. Modified date(First) : 2003/02/05
'*  7. Modified date(Last)  : 2003/02/06
'*  8. Modifier (First)     : Lee Woo Guen
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")

Dim oPB3S221

Dim strCode																	'☆ : Lookup 용 코드 저장 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strPlantCd
Dim GroupCount, GroupCount1
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 

Dim strClassCd
Dim strClassDesc
Dim strCharCd1
Dim strCharCd2
Dim strCharMgr

Dim strPrevNextFlg			'String
Dim I1_B_Class_Cd				'String
Dim E1_B_Class
Dim E2_B_Characteristic
Dim E3_B_Characteristic
Dim E4_StatusCodeOfPrevNext		'String

Const C_E1_Class_Cd = 0
Const C_E1_Class_Nm = 1
Const C_E1_Class_Digit = 2
Const C_E1_Class_Mgr = 3
Const C_E1_IsUsedByItem = 4
Const C_E2_Char_Cd = 0
Const C_E2_Char_Nm = 1
Const C_E2_Char_Value_Digit = 2
Const C_E3_Char_Cd = 0
Const C_E3_Char_Nm = 1
Const C_E3_Char_Value_Digit = 2

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next														'☜: 
Err.Clear																	'☜: Protect system from crashing

	strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
	strPrevNextFlg = Request("PrevNextFlg")
	strClassCd = Request("txtClassCd")
 
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)

	UNISqlId(0) = "b3b22mb1a"

	UNIValue(0, 0) = " " & FilterVar(UCase(strClassCd), "''", "S") & " "
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
	
	%>
	<Script Language=vbscript>
		parent.frm1.txtClassNm.value = ""
	</Script>
	<%

	' Class 명 Display      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("122650", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
		parent.frm1.txtClassCd.Focus()
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
		parent.frm1.txtClassNm.value = "<%=ConvSPChars(rs0("CLASS_NM"))%>"
		</Script>	
		<%
		rs0.Close
		Set rs0 = Nothing
	End If

	I1_B_Class_Cd = UCase(Trim(strClassCd))

	Set oPB3S221 = Server.CreateObject("PB3S221.cBLkupClassSvr")

    If CheckSYSTEMError(Err,True) = True Then
		Response.End 
    End If

	Call oPB3S221.B_LOOK_UP_CLASS_SVR(gStrGlobalCollection, _
								strPrevNextFlg, _
								I1_B_Class_Cd, _
								E1_B_Class, _
								E2_B_Characteristic, _
								E3_B_Characteristic, _
								E4_StatusCodeOfPrevNext)

    If CheckSYSTEMError(Err,True) = True Then
       Set oPB3S221 = Nothing
       Response.End 
    End If
    
    Set oPB3S221 = Nothing															'☜: Unload Comproxy
    
	If (E4_StatusCodeOfPrevNext = "900011" Or E4_StatusCodeOfPrevNext = "900012") Then
		Call DisplayMsgBox(E4_StatusCodeOfPrevNext, vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	End If
	
	Response.Write "<Script Language = VBScript> " & vbCr
	Response.Write "Dim LngRow " & vbCr
	
	If E1_B_Class(C_E1_IsUsedByItem) = True Then
		Response.Write "	parent.blnFlgIsUsedByItem = True" & vbCr
	Else
		Response.Write "	parent.blnFlgIsUsedByItem = False" & vbCr
	End If
		
	Response.Write "With parent.frm1 " & vbCr
	Response.Write "	.txtClassCd.value			= """ & ConvSPChars(UCase(Trim(E1_B_Class(C_E1_Class_Cd)))) & """" & vbCr
	Response.Write "	.txtClassNm.value			= """ & ConvSPChars(E1_B_Class(C_E1_Class_Nm)) & """" & vbCr
	Response.Write "	.txtClassCd1.value			= """ & ConvSPChars(UCase(Trim(E1_B_Class(C_E1_Class_Cd)))) & """" & vbCr
	Response.Write "	.txtClassNm1.value			= """ & ConvSPChars(E1_B_Class(C_E1_Class_Nm)) & """" & vbCr
	Response.Write "	.txtClassDigit.value		= """ & ConvSPChars(E1_B_Class(C_E1_Class_Digit)) & """" & vbCr
	Response.Write "	.cboClassMgr.value			= """ & ConvSPChars(E1_B_Class(C_E1_Class_Mgr)) & """" & vbCr
	
	Response.Write "	.txtCharCd1.value			= """ & ConvSPChars(UCase(Trim(E2_B_Characteristic(C_E2_Char_Cd)))) & """" & vbCr
	Response.Write "	.txtCharNm1.value			= """ & ConvSPChars(E2_B_Characteristic(C_E2_Char_Nm)) & """" & vbCr
	Response.Write "	.txtCharValueDigit1.value	= """ & ConvSPChars(E2_B_Characteristic(C_E2_Char_Value_Digit)) & """" & vbCr
	
	Response.Write "	.txtCharCd2.value			= """ & ConvSPChars(UCase(Trim(E3_B_Characteristic(C_E3_Char_Cd)))) & """" & vbCr
	Response.Write "	.txtCharNm2.value			= """ & ConvSPChars(E3_B_Characteristic(C_E3_Char_Nm)) & """" & vbCr
	Response.Write "	.txtCharValueDigit2.value	= """ & ConvSPChars(E3_B_Characteristic(C_E3_Char_Value_Digit)) & """" & vbCr
	
	Response.Write "	parent.lgNextNo = """"" & vbCr		' 다음 키 값 넘겨줌 
	Response.Write "	parent.lgPrevNo = """"" & vbCr		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 

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

