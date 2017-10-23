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
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         : MPS관리 
'*  5. Program Desc         : Query MPS
'*  6. Modified date(First) : 2000/11/02
'*  7. Modified date(Last)  : 2002/12/10
'*  8. Modifier (First)     : Lee Hyun Jae
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% 
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim i

Const C_SHEETMAXROWS_D = 500

Dim IntNextKey		' 다음 값 
Dim strItemCd
Dim strTrackingNo
Dim strMPSStatus
Dim lGetSvrDate

On Error Resume Next

Call HideStatusWnd

lGetSvrDate = GetSvrDate

strQryMode = Request("lgIntFlgMode")

	Err.Clear

	Redim UNISqlId(4)
	Redim UNIValue(4, 10)
	
	UNISqlId(0) = "p2217mb6"
	UNISqlId(1) = "184000saa"
	UNISqlId(2) = "184000sac"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "127400saa"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		strItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		strTrackingNo = FilterVar(Trim(Request("txtTrackingNo"))	, "''", "S")
	End IF

	IF Request("cboMPSStatus") = "" Then
		strMPSStatus = "|"
	Else
		strMPSStatus = FilterVar(Trim(Request("cboMPSStatus"))	, "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 2) = strItemCd
	UNIValue(0, 5) = strTrackingNo
	
	IF Request("txtPlndFromDt") = "" THEN
	   UNIValue(0, 3) = "|"
	ELSE
	   UNIValue(0, 3) = FilterVar(UniConvDate(Request("txtPlndFromDt"))	, "''", "S")
	END IF
	
	IF Request("txtPlndToDt") = "" THEN
	   UNIValue(0, 4) = "|"
	ELSE
	   UNIValue(0, 4) = FilterVar(UniConvDate(Request("txtPlndToDt"))	, "''", "S")
	END IF
	
	UNIValue(0, 6) = strMPSStatus
	UNIValue(0, 7) = "" & FilterVar("N", "''", "S") & " "
	UNIValue(0, 8) = "" & FilterVar("N", "''", "S") & " "
	IF Request("txtItemGroupCd") = "" Then
		UNIValue(0, 9) = "|"
	Else
		UNIValue(0, 9) = "d.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 10) = "|"
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 10) =  " (a.item_cd > " & FilterVar(Ucase(Request("lgStrPrevKey1")),"''","S")  & _
							" or ( a.item_cd = " & FilterVar(Ucase(Request("lgStrPrevKey1")),"''","S") & _
							" and a.tracking_no > " & FilterVar(Ucase(Request("lgStrPrevKey2")),"''","S") & _
							" ) or ( a.item_cd = " & FilterVar(Ucase(Request("lgStrPrevKey1")),"''","S") & _
							" and a.tracking_no = " & FilterVar(Ucase(Request("lgStrPrevKey2")),"''","S") & _
							" and a.mps_dt >= " & FilterVar(Ucase(Request("lgStrPrevKey3")),"''","S") & " )) "  
	End Select 

	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(2, 0) = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Trim(Request("txtItemGroupCd"))),"''","S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
	If Not(rs1.EOF AND rs1.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "With parent.frm1" & vbCrLf
				Response.Write ".txtPlantNm.value = """ & ConvSPChars(rs1("plant_nm")) & """" & vbCrLf
				Response.Write ".txtPH.text = """ & UniDateClientFormat(UniDateAdd("d", rs1("plan_hrzn"), lGetSvrDate, gServerDateFormat)) & """" & vbCrLf
				Response.Write ".txtDTF.text = """ & UniDateClientFormat(UniDateAdd("d", rs1("dtf_for_mps"), lGetSvrDate, gServerDateFormat)) & """" & vbCrLf
				Response.Write ".txtPTF.text = """ & UniDateClientFormat(UniDateAdd("d", rs1("ptf_for_mps"), lGetSvrDate, gServerDateFormat)) & """" & vbCrLf
			Response.Write "End With" & vbCrLf
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
		rs4.Close
		Set rs4 = Nothing
		Response.End
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim iTotalStr
Dim arrVal

With parent
	LngMaxRow = .frm1.vspdData.MaxRows
	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim arrVal(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim arrVal(<%=rs0.RecordCount - 1%>)
<%
		End If
		
		For i=0 to rs0.RecordCount-1		
			If i < C_SHEETMAXROWS_D Then 
%>

			strData = ""
			strData = strData & Chr(11) & "0"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("mps_dt"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("mps_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
			strData = strData & Chr(11) & "<%=UCase(Trim(rs0("mps_confirm_flg")))%>"			
			strData = strData & Chr(11) & "<%=UCase(Trim(rs0("mrp_confirm_flg")))%>"			

<%			If 	Trim(rs0("mps_status")) = "FM" Then %>
				strData = strData & Chr(11) & "Firm"
<%			ElseIf 	Trim(rs0("mps_status")) = "OP" Then %>
				strData = strData & Chr(11) & "Open"
<%			Else %>
				strData = strData & Chr(11) & "Plan"
<%			End if %>

			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("mps_no"))%>"	
			strData = strData & Chr(11) & "<%=UCase(rs0("prod_env"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("max_mrp_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("min_mrp_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("round_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=rs0("order_lt_mfg")%>"
			strData = strData & Chr(11) & "<%=rs0("order_lt_pur")%>"
			strData = strData & Chr(11) & "<%=rs0("mps_origin")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"	
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"	
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			ReDim Preserve arrVal(<%=i%>)
			arrVal(<%=i%>) = strData
<%		
			rs0.MoveNext
			End If
		Next
		
%>

		iTotalStr = Join(arrVal, "")
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey1 = "<%=Trim(rs0("item_cd"))%>"
		.lgStrPrevKey2 = "<%=Trim(rs0("tracking_no"))%>"
		.lgStrPrevKey3 = "<%=Trim(rs0("mps_dt"))%>"
		
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"  
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"   
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"                
		.frm1.hMPSOrigin.value		= "<%=Request("cboMPSOrigin")%>"
		.frm1.hMPSStatus.value		= "<%=Request("cboMPSStatus")%>"
		.frm1.hPlndFromDt.value		= "<%=Request("txtPlndFromDt")%>"
		.frm1.hPlndToDt.value		= "<%=Request("txtPlndToDt")%>"
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
		.DbQueryOk(LngMaxRow+1)
		
<%	End If
	
	rs0.Close
	Set rs0 = Nothing
	rs1.Close
	Set rs1 = Nothing
	rs2.Close
	Set rs2 = Nothing
	rs3.Close
	Set rs3 = Nothing
	rs4.Close
	Set rs4 = Nothing
%>
End With	
</Script>	
<%
Set ADF = Nothing
%>

<Script Language=vbscript RUNAT=server>

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>
