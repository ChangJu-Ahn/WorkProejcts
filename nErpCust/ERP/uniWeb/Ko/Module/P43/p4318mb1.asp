<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4318mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/08/25
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P","NOCOOKIE","MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim	rs0, rs1, rs2, rs3, rs4, rs5
Dim strQryMode
Dim strFlag
Dim lgStrPrevKey1
Dim strItemCd
Dim strWcCd
Dim strTrackingNo
Dim strItemGroupCd
Dim strItemAcct
Dim i	

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strMode = Request("txtMode")
	strQryMode = Request("lgIntFlgMode")

	On Error Resume Next

	lgStrPrevKey1 = UCase(Trim(Request("lgStrPrevKey1")))	
	
	'=======================================================================================================
	'	Handle Description
	'=======================================================================================================
	Redim UNISqlId(4)
	Redim UNIValue(4, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sac"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5)

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
	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
'			strFlag = "ERROR_ITEM"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs2.Close
			Set rs2 = Nothing
		End If
	Else
		rs2.Close
		Set rs2 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End IF
	' Work Center Check
	IF Request("txtWcCd") <> "" Then
	 	If rs3.EOF AND rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_WCCD"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtWCNm.value = """ & ConvSPChars(rs3("wc_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtWCNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
	' Tracking No. Check
	IF Request("txtTrackingNo") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			strFlag = "ERROR_TRACK"
		Else
			rs4.Close
			Set rs4 = Nothing
		End If
	Else
		rs4.Close
		Set rs4 = Nothing
	End IF
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs5.EOF AND rs5.BOF Then
			rs5.Close
			Set rs5 = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs5("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs5.Close
			Set rs5 = Nothing
		End If
	Else
		rs5.Close
		Set rs5 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
		
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtWCCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_TRACK" Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtTrackingNo.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		ElseIf strFlag = "ERROR_GROUP" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
		Set ADF = Nothing
		Response.End
	End IF
	Set ADF = Nothing
	'=======================================================================================================
	'	Main Query
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 8)
	
	UNISqlId(0) = "p4318mb1"
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & ""
	END IF
	
	IF Trim(lgStrPrevKey1) = "" Then
	   lgStrPrevKey1 = "|"
	ELSE
	   lgStrPrevKey1 = " " & FilterVar(UCase(lgStrPrevKey1), "''", "S") & ""
	END IF

	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = " " & FilterVar(UCase(Request("txtWcCd")), "''", "S") & ""
	END IF
		
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = " " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & ""
	END IF	
	
	IF Request("cboItemAcct") = "" Then
	   strItemAcct = "|"
	ELSE
	   strItemAcct = " " & FilterVar(UCase(Request("cboItemAcct") ), "''", "S") & ""
	END IF	

	IF Trim(Request("txtItemGroupCd")) = "" Then
	   strItemGroupCd = "|"
	ELSE
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 2) = strItemCd
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 2) = lgStrPrevKey1
	End Select
	
	UNIValue(0, 3) = strWcCd
	UNIValue(0, 4) = strTrackingNo
	UNIValue(0, 5) = strItemAcct
	UNIValue(0, 6) = FilterVar(UNIConvDate(Request("txtReqStartDt")), "''", "S")
	UNIValue(0, 7) = FilterVar(UNIConvDate(Request("txtReqEndDt")), "''", "S")
	UNIValue(0, 8) = strItemGroupCd
	
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
Dim TmpBuffer
Dim iTotalStr
    	
With parent
	LngMaxRow = .frm1.vspdData1.MaxRows
<%  
	If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
	End If

    For i=0 to rs0.RecordCount-1 
		If i < C_SHEETMAXROWS_D Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REQ_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RESVD_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("ISSUED_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("CONSUMED_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REMAIN_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_GROUP_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_GROUP_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_ACCT"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		End If
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey1 = "<%=Trim(rs0("ITEM_CD"))%>"		

		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hWcCd.value			= "<%=ConvSPChars(Request("txtWcCd"))%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		.frm1.hItemAcct.value		= "<%=ConvSPChars(Request("cboItemAcct"))%>"
		.frm1.hReqStartDt.value		= "<%=Request("txtReqStartDt")%>"
		.frm1.hReqEndDt.value		= "<%=Request("txtReqEndDt")%>"
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
