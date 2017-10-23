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
'*  3. Program ID           : p2214mb1.asp
'*  4. Program Name         : MPS History
'*  5. Program Desc         : 
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2003/02/24
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% 

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1
Dim a


Dim lgStrPrevKey	' 이전 값 

Call HideStatusWnd

On Error Resume Next

Dim strRunNo
	lgStrPrevKey = Request("lgStrPrevKey")
	
	Err.Clear
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)
	
	UNISqlId(0) = "184300saa"
	UNISqlId(1) = "184000saa"
	
	IF Request("txtRunNO") = "" Then
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
		Response.Write "	parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("plant_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "	parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If
	
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing

		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim strStartDt
Dim arrVal
ReDim arrVal(0)
    	
With parent
	LngMaxRow = .frm1.vspdData.MaxRows
		
<%  For a=0 to rs0.RecordCount-1 %>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MPS_HISTORY_NO"))%>"

<%		IF 	rs0("STATUS") = "1" THEN%>
			strData = strData & Chr(11) & "일괄생성"
<% 		ELSEIF rs0("STATUS") = "2" THEN%>
			strData = strData & Chr(11) & "승인"
<% 		ELSEIF rs0("STATUS") = "3" THEN%>
			strData = strData & Chr(11) & "일괄생성취소"
<%		END IF%>
		
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("START_DT")) & "  " & FormatDateTime(rs0("START_DT"),3)%>"	'MPS실행일 
		
<%		IF 	rs0("START_FLG") = "D" THEN%>
			strData = strData & Chr(11) & "<%=UNIDateAdd("d",UniCInt(rs0("mps_dtf"), 0), UniDateClientFormat(rs0("start_dt")), gDateFormat)%>"   	'기준일자 
<%		ELSE%>
			strData = strData & Chr(11) & "<%=UNIDateAdd("d",UniCInt(rs0("mps_ptf"), 0), UniDateClientFormat(rs0("start_dt")), gDateFormat)%>"   	
<%      END IF%>

		strData = strData & Chr(11) & "<%=UNIDateAdd("d",UniCInt(rs0("mps_dtf"), 0), UniDateClientFormat(rs0("start_dt")), gDateFormat)%>"   	'DTF
		strData = strData & Chr(11) & "<%=UNIDateAdd("d",UniCInt(rs0("mps_ptf"), 0), UniDateClientFormat(rs0("start_dt")), gDateFormat)%>"   	'PTF
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("PLAN_DT")) %>"			'계획일자 
		strData = strData & Chr(11) & "<%=rs0("INV_FLG")%>"				'가용재고 
		strData = strData & Chr(11) & "<%=rs0("SS_FLG")%>"				'안전재고 
		strData = strData & Chr(11) & "<%=rs0("MAX_ORDER_FLG")%>"		'최대Lot
		strData = strData & Chr(11) & "<%=rs0("MIN_ORDER_FLG")%>"		'최소Lot
		strData = strData & Chr(11) & "<%=rs0("ROUND_FLG")%>"			'올림수 
		strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("APPROVE_DT")) & "  " & FormatDateTime(rs0("APPROVE_DT"),3)%>"	'승인일 
		strData = strData & Chr(11) & "<%=UCase(ConvSPChars(rs0("APPROVER")))%>"				'승인자 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("START_NO"))%>"				'시작MPS No
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("END_NO"))%>"					'종료MPS No
		strData = strData & Chr(11) & "<%=rs0("START_DT_HD")%>"					'MPS실행일(정렬을 위한)
		strData = strData & Chr(11) & "<%=rs0("APPROVE_DT_HD")%>"					'승인일(정렬을 위한)
		
		strData = strData & Chr(11) & LngMaxRow + <%=a%>
		strData = strData & Chr(11) & Chr(12)
		
		ReDim Preserve arrVal(<%=a%>)
		arrVal(<%=a%>) = strData
<%		
		rs0.MoveNext
	Next
%>
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData Join(arrVal,"")

<%			
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing
%>
