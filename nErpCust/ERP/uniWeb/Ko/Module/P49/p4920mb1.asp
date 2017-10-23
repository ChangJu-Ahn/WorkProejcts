<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4911mb2.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005-01-25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag

Dim pPP4G920

Dim lgIntFlgMode

Dim iErrorPosition
Dim strPlantCd
Dim strWorkCd
Dim strResourceCd
'Dim Err
Dim rs1, rs2, rs3, rs4

strPlantCd    = UCase(Trim(Request("txtPlantCd")))
strWorkCd     = UCase(Trim(Request("txtWorkDt")))
strResourceCd = UCase(Trim(Request("txtResourceCd")))

'Response.Write strPlantCd  & "<br>"

	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000san"

	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtResourceCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)

'	Response.Write strRetMsg & "<P>"
'	Response.Write UNIValue(0, 0) & "<P>"
'	Response.Write UNIValue(1, 0) & "<P>"
'	Response.Write UNIValue(1, 1) & "<P>"

	' Plant 명 Display
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs1.Close
		Set rs1 = Nothing
	End If

	' 자원명 Display
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing
		Call DisplayMsgBox("181600", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtResourceNm.value = """"" & vbcr
		Response.Write "parent.frm1.txtResourceCd.focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtResourceNm.value = """ & ConvSPChars(rs2("description")) & """" & vbcr
		Response.Write "</Script>" & vbcr
	End If
'-------------------------------------------------------------------------------
'	COM+ Action
'-------------------------------------------------------------------------------
'On Error Resume Next
'Err.Clear

Set pPP4G920 = Server.CreateObject("PP4G920.cPMngWorkTimeBatch")

If CheckSYSTEMError(Err,True) = True Then
%>
<Script Language=vbscript>
With parent
	.DbSaveFail()
End With
</Script>
<%
   Response.End
End If

Call pPP4G920.P_MANAGE_WORK_TIME_BATCH(gStrGlobalCollection, strPlantCd, strWorkCd, strResourceCd)

If CheckSYSTEMError2(Err, True, "", "", "", "", "") = True Then
	Set pPP1G303 = Nothing	
	Response.End
End If

Set pPP4G920 = Nothing

%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>
<%
' Server Side 로직은 여기서 끝남 
'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>
