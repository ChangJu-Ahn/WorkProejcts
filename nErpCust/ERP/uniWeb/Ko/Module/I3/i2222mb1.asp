<%'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2222mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2005/01/25
'*  8. Modifier (First)     : Chen, Jae Hyun
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
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim	rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim strQryMode
Dim strFlag
Dim lgStrPrevKey1
Dim strItemCd
Dim strSLCd
Dim strTrackingNo
Dim strItemGroupCd
Dim strIssueFlag
Dim strBaseUnit
Dim strGoodQty
Dim strSign
Dim strPrevKey1, strPrevKey2, strPrevKey3
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
	Redim UNISqlId(5)
	Redim UNIValue(5, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sad"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "180000sas"
	UNISqlId(5) = "s0000qa003"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtSlCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")
	UNIValue(5, 0) = Filtervar(Ucase(Request("txtBaseUnit")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5, rs6)

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
	' S/L Check
	IF Request("txtSlCd") <> "" Then
	 	If rs3.EOF AND rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_SLCD"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """ & ConvSPChars(rs3("sl_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf
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
	' Unit Check
	IF Request("txtBaseUnit") <> "" Then
	 	If rs6.EOF AND rs6.BOF Then
			rs6.Close
			Set rs6 = Nothing
			strFlag = "ERROR_UNIT"
		Else
			rs6.Close
			Set rs6 = Nothing
		End If
	Else
		rs6.Close
		Set rs6 = Nothing
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
		ElseIf strFlag = "ERROR_SLCD" Then
			Call DisplayMsgBox("125700", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSLCd.focus" & vbCrLf
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
		ElseIf strFlag = "ERROR_UNIT" Then
			Call DisplayMsgBox("124000", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtUnit.focus" & vbCrLf
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
	Redim UNIValue(0, 9)
	
	UNISqlId(0) = "i2222mb1"
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtSLCd")) = "" Then
	   strSLCd = "|"
	ELSE
	   strSLCd = FilterVar(UCase(Request("txtSLCd")), "''", "S")
	END IF
		
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	END IF	
	
	IF Request("rdoIssueFlag") = "" Then
	   strIssueFlag = "|"
	ELSE
	   strIssueFlag = " A.SCHD_ISSUE_QTY > 0 "
	END IF	

	IF Trim(Request("txtItemGroupCd")) = "" Then
	   strItemGroupCd = "|"
	ELSE
		strItemGroupCd = " B.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	END IF
	
	IF Trim(Request("txtBaseUnit")) = "" Then
	   strBaseUnit = "|"
	ELSE
		strBaseUnit = FilterVar(UCase(Request("txtBaseUnit") ), "''", "S")
	END IF
	
	StrPrevKey1 = FilterVar(Request("lgStrPrevKey1"), "''", "S")
	StrPrevKey2 = FilterVar(Request("lgStrPrevKey2"), "''", "S")
	StrPrevKey3 = FilterVar(Request("lgStrPrevKey3"), "''", "S")
	
	strSign = Trim(Request("txtSign"))
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strSLCd
	UNIValue(0, 3) = strTrackingNo
	UNIValue(0, 4) = strBaseUnit
	UNIValue(0, 5) = strItemCd
	UNIValue(0, 6) = strIssueFlag
	UNIValue(0, 7) = strItemGroupCd 
	
	
	Select Case Request("rdoInventoryFlag")		
		Case "Y"
			UNIValue(0, 8) = " A.GOOD_ON_HAND_QTY " & strSign &  UniConvNum(Request("txtBaseQty"), 0)
		Case "N"
			UNIValue(0, 8) = " A.PREV_GOOD_QTY "  & strSign &  UniConvNum(Request("txtBaseQty"), 0)
		Case Else
			UNIValue(0, 8) = " (A.GOOD_ON_HAND_QTY "  & strSign &  UniConvNum(Request("txtBaseQty"), 0) & _
							" Or A.PREV_GOOD_QTY "  & strSign & UniConvNum(Request("txtBaseQty"), 0) & ")"
	End Select	

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 9) =  "|"
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 9) = " (A.ITEM_CD > " & StrPrevKey1 & _
							 " OR (A.ITEM_CD = " & strPrevKey1 & " AND A.SL_CD > " & strPrevKey2 & _
							 " ) OR ( A.ITEM_CD = " & strPrevKey1 & " AND A.SL_CD = " & strPrevKey2 & " AND A.TRACKING_NO >= "  & strPrevKey3 & "))"					 
	End Select
	
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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BLOCK_INDICATOR"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASIC_UNIT"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_ON_HAND_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_ON_HAND_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_ON_INSP_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("STK_IN_TRNS_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("SCHD_RCPT_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("SCHD_ISSUE_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PREV_GOOD_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PREV_BAD_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PREV_STK_IN_TRNS_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("ALLOCATION_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat("0",ggQty.DecPoint,0)%>"
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
		.lgStrPrevKey2 = "<%=Trim(rs0("SL_CD"))%>"	
		.lgStrPrevKey3 = "<%=Trim(rs0("TRACKING_NO"))%>"

		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hSLCd.value			= "<%=ConvSPChars(Request("txtSLCd"))%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		.frm1.hBaseQty.value		= "<%=ConvSPChars(Request("txtBaseQty"))%>"
		.frm1.hSchedIssueFlg.value	= "<%=ConvSPChars(Request("rdoIssueFlag"))%>"
		.frm1.hInventoryFlg.value	= "<%=ConvSPChars(Request("rdoInventoryFlag"))%>"
		.frm1.hBaseUnit.value		= "<%=ConvSPChars(Request("txtBaseUnit"))%>"
		.frm1.hcboCompareFlag.value =  "<%=ConvSPChars(Request("txtSign"))%>"
		
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
