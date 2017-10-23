<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<% Call LoadBasisGlobalInf
'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/15
'*  7. Modified date(Last)  : 2000/09/26
'*  8. Modifier (First)     : Lee Hyun Jae
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
Dim strQryMode
Dim i

Dim lgStrPrevKey
Dim strItemCd

Call HideStatusWnd



On Error Resume Next


	Redim UNISqlId(0)
	Redim UNIValue(0, 3)
	
	UNISqlId(0) = "185500saa"

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("lgPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("lgMrpRunNo")), "''", "S")
	
	IF Request("lgItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = " " & FilterVar(UCase(Request("lgItemCd")), "''", "S") & ""
	END IF
	UNIValue(0, 3) = strItemCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		
		rs0.Close
		Set rs0 = Nothing
					
		Response.End
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim arrVal
ReDim arrVal(0)
    	
With parent
	LngMaxRow = .vspdData.MaxRows
	
<%    For i=0 to rs0.RecordCount-1 %>

			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ERROR_NM"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("START_PLAN_DT"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("END_PLAN_DT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLAN_ORDER_NO"))%>"
			strData = strData & Chr(11) & LngMaxRow + "<%=i%>"
			strData = strData & Chr(11) & Chr(12)
			
			ReDim Preserve arrVal(<%=i%>)
			arrVal(<%=i%>) = strData

<%		
			rs0.MoveNext
	Next	
%>
	.ggoSpread.Source = .vspdData 
	.ggoSpread.SSShowData Join(arrVal,"")
	
.DbQueryOk
End With
</Script>	
<%
rs0.Close
Set rs0 = Nothing

Set ADF = Nothing
%>
