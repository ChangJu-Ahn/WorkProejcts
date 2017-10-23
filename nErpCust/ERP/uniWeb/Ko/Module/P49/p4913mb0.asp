<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4913mb1.asp
'*  4. Program Name         : 작업일보 등록
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005-01-17
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Yoon, Jeong Woo
'*  9. Modifier (Last)      :
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim lgStrPrevKey
Dim strProdOrdNo
Dim strFlag
Dim strWcCd
Dim i

	Call HideStatusWnd

	'=======================================================================================================
	'	Main Query - Order Header Display
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	UNISqlId(0) = "P4913MA0"

'	UNIValue(0, 0) = "^"
	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 1) = FilterVar(UNIConvDate(Request("txtprodDt")),"''","S")
	UNIValue(0, 2) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
'		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent.frm1" & vbCrLf
Response.Write "	.fpDoubleSingle1.value = """ & rs0("JK_MAN") & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle2.value = """ & rs0("INC_MAN") & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle3.value = """ & rs0("OT_MAN_CNT") & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle4.value = """ & ConvToTimeFormat(rs0("WK_TIME")) & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle5.value = """ & rs0("JK_TIME") & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle6.value = """ & ConvToTimeFormat(rs0("INC_TIME")) & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle7.value = """ & ConvToTimeFormat(rs0("OT_MAN_TIME")) & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle8.value = """ & ConvToTimeFormat(rs0("WK_LOSS_TIME")) & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle9.value = """ & rs0("HOLIDAY_MAN") & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle10.value = """ & rs0("DESC_MAN") & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle11.value = """ & ConvToTimeFormat(rs0("OT_MAN_TOTAL")) & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle12.value = """ & rs0("HOLIDAY_TIME") & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle13.value = """ & ConvToTimeFormat(rs0("DESC_TIME")) & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle14.value = """ & ConvToTimeFormat(rs0("ETC_TIME")) & """" & vbCrLf		'☆: 
Response.Write "	.fpDoubleSingle15.value = """ & ConvToTimeFormat(rs0("WK_REAL_TIME")) & """" & vbCrLf		'☆: 

Response.Write "	parent.DbQueryOkForm(0)" & vbCrLf																'☜: 조화가 성공 
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																	'☜: Process End
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : 시간 형식으로 변경 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec

	Dim iVal2

	iVal2 = Fix(iVal)
	
	If iVal2 = 0 Then
		ConvToTimeFormat = "00:00:00"
	ElseIf iVal2 > 0 Then
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)

		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
	Else
		iVal2 = Replace(iVal2, "-", "")
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		ConvToTimeFormat = "-" & ConvToTimeFormat
	End If
End Function
</script>