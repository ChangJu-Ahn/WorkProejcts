<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2347mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")  

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1

Dim lgStrPrevKey	' 이전 값 
Dim i

On Error Resume Next

Dim strRunNo
	lgStrPrevKey = Request("lgStrPrevKey")
	
	Err.Clear
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "185700saa"
	UNISqlId(1) = "184000saa"
	
	IF Request("txtRunNo") = "" Then
		strRunNo = "|"
	Else
		StrRunNo = FilterVar(Trim(Request("txtRunNo"))	, "''", "S")
	End IF
		
	UNIValue(0, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 1) = strRunNo
	
	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    If Not(rs1.EOF AND rs1.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("plant_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If
    
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing			
		Response.End
		
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim arrVal
ReDim arrVal(0)
    	
With parent
	LngMaxRow = .frm1.vspdData.MaxRows
		
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RUN_NO"))%>"

<%		IF 	rs0("STATUS") = "1" THEN
%>			strData = strData & Chr(11) & "전개"
		
<% 		ELSEIF rs0("STATUS") = "2" THEN
%>			strData = strData & Chr(11) & "승인"

<% 		ELSEIF rs0("STATUS") = "3" THEN
%>			strData = strData & Chr(11) & "부분전환"

<% 		ELSEIF rs0("STATUS") = "4" THEN
%>			strData = strData & Chr(11) & "전환완료"

<% 		ELSEIF rs0("STATUS") = "5" THEN
%>			strData = strData & Chr(11) & "전개취소"

<% 		ELSE
%>			strData = strData & Chr(11) & "승인취소"

<%		END IF
%>		
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("START_DT")) & "  " & FormatDateTime(rs0("START_DT"),3)%>"
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("APPROVE_DT")) & "  " & FormatDateTime(rs0("APPROVE_DT"),3)%>"
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("CONV_DT")) & "  " & FormatDateTime(rs0("CONV_DT"),3)%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("from_during_dt")) %>"   		 
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("FIRM_DURING_DT"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("TO_DURING_DT"))%>"
		strData = strData & Chr(11) & "<%=rs0("INV_FLG")%>"
		strData = strData & Chr(11) & "<%=rs0("SS_FLG")%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("APPROVER"))%>"
		strData = strData & Chr(11) & "<%=rs0("ORDER_QTY")%>"
		strData = strData & Chr(11) & "<%=rs0("CONV_ORDER_QTY")%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("START_ORDER_NO"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("END_ORDER_NO"))%>"
		strData = strData & Chr(11) & "<%=rs0("START_DT_HD")%>"
		strData = strData & Chr(11) & "<%=rs0("APPROVE_DT_HD")%>"
		strData = strData & Chr(11) & "<%=rs0("CONV_DT_HD")%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		ReDim Preserve arrVal(<%=i%>)
		arrVal(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData Join(arrVal,"")
		
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<% 
Set ADF = Nothing
%>
