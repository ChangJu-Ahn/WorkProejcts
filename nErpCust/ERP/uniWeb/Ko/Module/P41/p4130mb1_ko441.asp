<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4213mb1.asp
'*  4. Program Name         : Cancel Order Release
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/03/30
'*  7. Modified date(Last)  : 2002/07/18
'*  8. Modifier (First)     : Park, Bumsoo
'*  9. Modifier (Last)      : Chen, Jae Hyun
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
Dim strItemCd
Dim strProdOrdNo
Dim strTrackingNo
Dim strOrderType
Dim strItemGroupCd
Dim strFlag
Dim strSendStatus  '20080115::hanc

Dim i

	Const C_SHEETMAXROWS_D = 100

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

	lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")
	
	'======================================================================================================
	'	Handle Description
	'======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)
	
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sam"
	UNISqlId(3) = "180000sas"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtItemGroupCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	
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
			strFlag = "ERROR_ITEM"
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
	' Tracking No. Check
	IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_TRACK"
		Else
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		rs3.Close
		Set rs3 = Nothing
	End IF
	' Item Group Check
	IF Request("txtItemGroupCd") <> "" Then
	 	If rs4.EOF AND rs4.BOF Then
			rs4.Close
			Set rs4 = Nothing
			strFlag = "ERROR_GROUP"
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs4("item_group_nm")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			rs4.Close
			Set rs4 = Nothing
		End If
	Else
		rs4.Close
		Set rs4 = Nothing
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
	'	Main Query - Order Header Display
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 8)
	
	UNISqlId(0) = "P4130MB1_KO441"  '20080115::hanc
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF
		
	IF Trim(Request("txtProdOrderNo")) = "" Then
	   strProdOrdNo = "|"
	ELSE
	   strProdOrdNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	END IF
	
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	END IF	
	
	'20080115::hanc
	IF Request("rdoSendStatus") = "" Then
		strSendStatus = "|"
	Else
		strSendStatus = FilterVar(UCase(Request("rdoSendStatus")), "''", "S")
	End If

	IF Request("cboOrderType") = "" Then
	   strOrderType = "|"
	ELSE
	   strOrderType = FilterVar(UCase(Request("cboOrderType")), "''", "S")
	END IF	
	
	IF Request("txtItemGroupCd") = "" Then
		strItemGroupCd = "|"
	Else
		strItemGroupCd = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 2) = strProdOrdNo
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 2) = lgStrPrevKey
	End Select
	UNIValue(0, 3) = strItemCd
	UNIValue(0, 4) = FilterVar(UNIConvDate(Request("txtProdFromDt")), "''", "S")
	UNIValue(0, 5) = FilterVar(UNIConvDate(Request("txtProdToDt")), "''", "S")
	UNIValue(0, 6) = strTrackingNo	
	UNIValue(0, 7) = strSendStatus      '20080115::hanc
	UNIValue(0, 8) = strOrderType
	UNIValue(0, 9) = strItemGroupCd
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow		
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
			strData = strData & Chr(11) & ""		
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("IF_FLAG_NM"))%>"	        '20080115::hanc	
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"		
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ROUT_NO"))%>"	
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("PLAN_START_DT"))%>"		
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("PLAN_COMPT_DT"))%>"		
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PRODT_ORDER_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
			strData = strData & Chr(11) & "<%=rs0("PRODT_ORDER_TYPE")%>"
			strData = strData & Chr(11) & "<%=rs0("PRODT_ORDER_TYPE")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
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
		
		.lgStrPrevKey1 = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"		

		.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hProdOrderNo.value= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
		.frm1.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hProdFromDt.value = "<%=Request("txtProdFromDt")%>"
		.frm1.hProdToDt.value	= "<%=Request("txtProdToDt")%>"		
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"		
		.frm1.hSendStatus.value= "<%=ConvSPChars(Request("rdoSendStatus"))%>"
		.frm1.hOrderType.value	= "<%=ConvSPChars(Request("cboOrderType"))%>"
		.frm1.hItemGroupCd.value= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"		
		
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbQueryOk(LngMaxRow)
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
