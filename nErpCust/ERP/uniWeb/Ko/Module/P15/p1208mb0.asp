<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1208mb0.asp
'*  4. Program Name         : Look up Manufacturing Instruction	Detail
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002/03/25
'*  7. Modified date(Last)  : 2002/11/20
'*  8. Modifier (First)     : Jeon, Jaehyun
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0										'DBAgent Parameter 선언 

Dim iLngRow

Dim StrWICd
Dim strWIDesc
Dim strValidStartDt
Dim strValidEndDt

' Order information Display
Redim UNISqlId(0)
Redim UNIValue(0, 1)

iLngRow = Request("txtRow")

UNISqlId(0) = "P1208MB0"
	
UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(UCase(Request("txtWICd")), "''", "S")
UNIValue(0, 2) = " " & FilterVar(UniConvDate(Request("txtStdDt")), "''", "S") & ""
UNIValue(0, 3) = " " & FilterVar(UniConvDate(Request("txtStdDt")), "''", "S") & ""
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("181423", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "Call parent.LookupWIFail(""" & iLngRow & """)" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End													'☜: 비지니스 로직 처리를 종료함 
End If
	
StrWICd = Trim(ConvSPChars(rs0("MFG_INSTRUCTION_DTL_CD")))
strWIDesc = Trim(ConvSPChars(rs0("MFG_INSTRUCTION_DTL_DESC")))
strValidStartDt = UNIDateClientFormat(rs0("VALID_START_DT"))
strValidEndDt = UNIDateClientFormat(rs0("VALID_END_DT"))
	
rs0.Close
Set rs0 = Nothing

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "Call parent.LookupWISuccess(""" & StrWICd & """, """ & strWIDesc & """, """ & strValidStartDt & """, """ & strValidEndDt & """, """ & iLngRow & """)" & vbCrLf
Response.Write "</Script>" & vbCrLf	
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
