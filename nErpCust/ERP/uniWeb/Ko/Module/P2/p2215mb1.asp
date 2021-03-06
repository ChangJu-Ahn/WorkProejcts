<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/02
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================-->
<% 

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim iIntCnt

Const C_SHEETMAXROWS = 100

Dim IntNextKey		' ���� �� 
Dim strItemCd
Dim strTrackingNo
Dim strReqStatus
Dim lgStrPrevKey1
Dim lgStrPrevKey2

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear                                      

	lgStrPrevKey1 = Trim(Request("lgStrPrevKey1"))
	lgStrPrevKey2 = Request("lgStrPrevKey2")
	
	Redim UNISqlId(4)
	Redim UNIValue(4, 9)
	
	UNISqlId(0) = "P2215MB1"
	UNISqlId(1) = "184000saa"
	UNISqlId(2) = "184000sac"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "127400saa"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(Trim(Request("txtTrackingNo"))	, "''", "S")
	End IF
	
	IF Request("cboReqStatus") = "" Then
		strReqStatus = "|"
	Else
		strReqStatus = FilterVar(Trim(Request("cboReqStatus"))	, "''", "S")
	End IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")

	Select Case strQryMode
	Case CStr(OPMD_CMODE)
		UNIValue(0, 2) = strItemCd
	Case CStr(OPMD_UMODE)
		UNIValue(0, 2) = FilterVar(lgStrPrevKey1 , "''", "S")

	End Select
	UNIValue(0, 3) = strTrackingNo		
	IF Request("txtFromReqrdDt") = "" THEN
	   UNIValue(0, 4) = "|"
	ELSE
	   UNIValue(0, 4) = FilterVar(UniConvDate(Request("txtFromReqrdDt")), "''", "S")

	END IF

    IF Request("txtToReqrdDt") = "" THEN
    	UNIValue(0, 5) = "|"
    ELSE
    	UNIValue(0, 5) = FilterVar(UniConvDate(Request("txtToReqrdDt")), "''", "S")

    END IF
	UNIValue(0, 6) = strReqStatus	
	Select Case strQryMode
	Case CStr(OPMD_CMODE)
		UNIValue(0, 7) = "|"
		UNIValue(0, 8) = "|"
	Case CStr(OPMD_UMODE)		
		UNIValue(0, 7) = "a.item_cd > " & FilterVar(lgStrPrevKey1 , "''", "S") & " or (a.item_cd = " & FilterVar(lgStrPrevKey1	, "''", "S")
		UNIValue(0, 8) = FilterVar(lgStrPrevKey2 , "''", "S")
	End Select
	
	IF Request("txtItemGroupCd") = "" Then
		UNIValue(0, 9) = "|"
	Else
		UNIValue(0, 9) = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	UNIValue(1, 0) = FilterVar(Request("txtPlantCd")	, "''", "S")
	UNIValue(2, 0) = FilterVar(Request("txtItemCd")	, "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")),"''","S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
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

	If Not(rs2.EOF AND rs2.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("item_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
	
	IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF AND rs3.BOF) Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtTrackingNo.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		End If
	End If
	
	If Not(rs4.EOF AND rs4.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs4("item_group_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf

		IF Request("txtItemGroupCd") <> "" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End 
		End If
	End If

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing	
		rs1.Close
		Set rs1 = Nothing					
		rs2.Close
		Set rs2 = Nothing		
		rs3.Close
		Set rs3 = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim arrVal
Redim arrVal(0)
    	
With parent
	LngMaxRow = .frm1.vspdData.MaxRows
<%  
    For iIntCnt = 0 to rs0.RecordCount-1 
		IF iIntCnt < C_SHEETMAXROWS THEN
%>  
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("req_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Issued_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("req_dt"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("req_type"))%>"

<%			if 	Trim(rs0("status")) = "AC" then	%>
				strData = strData & Chr(11) & "Accepted"
<%			else %>
				strData = strData & Chr(11) & "Requested"
<%			end if %>

			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("so_no"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("so_dtl_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"						
			strData = strData & Chr(11) & LngMaxRow + <%=iIntCnt%>
			strData = strData & Chr(11) & Chr(12)
			
			ReDim Preserve arrVal(<%=iIntCnt%>)
			arrVal(<%=iIntCnt%>) = strData			
<%		
			rs0.MoveNext
		END IF
	Next
%>
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData Join(arrVal,"")
	
		.lgStrPrevKey1 = "<%=rs0("ITEM_CD")%>"
		.lgStrPrevKey2 = "<%=rs0("ind_req_no")%>"

		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hReqStatus.value		= "<%=ConvSPChars(Request("cboReqStatus"))%>"
		.frm1.hFromReqrdDt.value	= "<%=Request("txtFromReqrdDt")%>"
		.frm1.hToReqrdDt.value		= "<%=Request("txtToReqrdDt")%>"
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
<%			
		rs0.Close
		Set rs0 = Nothing
		rs1.Close
		Set rs1 = Nothing		
		rs2.Close
		Set rs2 = Nothing			
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing
%>
