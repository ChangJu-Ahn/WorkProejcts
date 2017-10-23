<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!--'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2213pb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/12/10
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% 

Call LoadBasisGlobalInf

On Error Resume Next
Err.Clear

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
Dim i
Dim strRunNo
Dim LngMaxRow, strData
Dim arrVal
ReDim arrVal(0)

LngMaxRow = Request("txtMaxRows")

Call HideStatusWnd
	

	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "184300saa"
	
	IF Request("txtRunNO") = "" Then
		strRunNo = "|"
	Else
		StrRunNo = FilterVar(Trim(Request("txtRunNo"))	, "''", "S")
	End IF
		
	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = strRunNo
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With Parent	" & vbCrLf
		
	For i=0 to rs0.RecordCount-1 
		strData = ""
		strData = strData & Chr(11) & ConvSPChars(rs0("MPS_HISTORY_NO")) 
		strData = strData & Chr(11) & UNIDateClientFormat(rs0("START_DT"))
		strData = strData & Chr(11) & LngMaxRow + i 
		strData = strData & Chr(11) & Chr(12)
		
		ReDim Preserve arrVal(i)
		arrVal(i) = strData
			
		rs0.MoveNext
	Next

	Response.Write "	.ggoSpread.Source = .vspdData" & vbCrLf
	Response.Write "	.ggoSpread.SSShowData""" & Join(arrVal,"") & """" & vbCrLf
		
	rs0.Close
	Set rs0 = Nothing
		
	Response.Write "	.DbQueryOk" & vbCrLf
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf		

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

