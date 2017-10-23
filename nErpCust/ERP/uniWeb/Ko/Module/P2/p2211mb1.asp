<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2211mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Hyun Jae
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

On Error Resume Next								'☜: 

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4
Dim strQryMode
Dim i

Const C_SHEETMAXROWS = 100

Dim lgStrPrevKey	' 이전 값 
Dim lgStrPrevKey2	' 이전 값 (Tracking_No)
Dim lgStrPrevKey3	' 이전 값 (Due_Dt)
Dim lgStrPrevKey4	' 이전 값 (Split_Seq_No)

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strItemCd
Dim strTrackingNo
Dim strFromPlanDt
Dim strToPlanDt

	lgStrPrevKey = FilterVar(Request("lgStrPrevKey"), "''", "S")
	
	Redim UNISqlId(4)
	Redim UNIValue(4, 5)
	
	UNISqlId(0) = "p2211mb1"				' Data Query
	UNISqlId(1) = "184000saa"				' Plant Check
	UNISqlId(2) = "184000sac"				' Item Check
	UNISqlId(3) = "180000sam"				' Tracking No
	UNISqlId(4) = "127400saa"
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(Request("txtItemCd")	, "''", "S")
	END IF
	
	IF Request("txtTrackingNo") = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(Request("txtTrackingNo")	, "''", "S")
	END IF
	
	IF Request("txtFromPlanDt") = "" THEN
	   strFromPlanDt = "|"
	ELSE
	   strFromPlanDt = FilterVar(UniConvDate(Request("txtFromPlanDt"))	, "''", "S")
	END IF

    IF Request("txtToPlanDt") = "" THEN
    	strToPlanDt = "|"
    ELSE
    	strToPlanDt = FilterVar(UniConvDate(Request("txtToPlanDt"))	, "''", "S")
    END IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Request("txtPlantCd")	, "''", "S")

	Select Case strQryMode	
		Case CStr(OPMD_CMODE)
			If strItemCd = "|" and strTrackingNo = "|" Then             ' 품목, Tracking No 둘다 공백 
				UNIValue(0, 2) = "|"
			ElseIf strItemCd <> "|" and strTrackingNo = "|" Then        ' 품목만 입력 
				UNIValue(0, 2) = "a.item_cd >= " & strItemCd
			ElseIf strItemCd = "|" and strTrackingNo <> "|" Then        ' Tracking No만 입력 
				UNIValue(0, 2) = "a.tracking_no = " & strTrackingNo
			Else                                                        ' 품목, Tracking No 둘다 입력 
				UNIValue(0, 2) = "(a.item_cd >= " & strItemCd & " and a.tracking_no = " & strTrackingNo & ")"
			End If			
		Case CStr(OPMD_UMODE)
			lgStrPrevKey2 = FilterVar(Request("lgStrPrevKey2"), "''", "S")
			lgStrPrevKey3 = FilterVar(Request("lgStrPrevKey3"), "''", "S")
			lgStrPrevKey4 = FilterVar(Request("lgStrPrevKey4"), "''", "S")

			If strTrackingNo = "|" Then
				UNIValue(0, 2) = "((a.item_cd = " & lgStrPrevKey & " and a.tracking_no >= " & lgStrPrevKey2 & _
							" and a.due_dt = " & lgStrPrevKey3 & " and splt_seq_no >= " & lgStrPrevKey4 & _
							") or (a.item_cd >= " & lgStrPrevKey & " and a.due_dt > " & lgStrPrevKey3 & _
							" and splt_seq_no >= " & lgStrPrevKey4 & ") or (a.item_cd > " & lgStrPrevKey & "))"		
			Else
				UNIValue(0, 2) = "((a.item_cd = " & lgStrPrevKey & " and a.tracking_no = " & lgStrPrevKey2 & _
							" and a.due_dt = " & lgStrPrevKey3 & " and splt_seq_no >= " & lgStrPrevKey4 & _
							") or (a.item_cd >= " & lgStrPrevKey & " and a.tracking_no = " & strTrackingNo & _
							" and a.due_dt > " & lgStrPrevKey3 & " and splt_seq_no >= " & lgStrPrevKey4 & _
							") or (a.item_cd > " & lgStrPrevKey & " and a.tracking_no = " & strTrackingNo & "))"
			End If

	End Select	
	
	UNIValue(0, 3) = strFromPlanDt
	UNIValue(0, 4) = strToPlanDt
	IF Request("txtItemGroupCd") = "" Then
		UNIValue(0,5) = "|"
	Else
		UNIValue(0,5) = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF

	UNIValue(1, 0) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(2, 0) = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Trim(Request("txtItemGroupCd"))),"''","S")
	
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
		rs4.Close
		Set rs4 = Nothing
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
		If i < C_SHEETMAXROWS Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("due_dt"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
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
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData Join(arrVal,"")
		
		.lgStrPrevKey = "<%=rs0("item_cd")%>"
		.lgStrPrevKey2 = "<%=rs0("tracking_no")%>"
		.lgStrPrevKey3 = "<%=rs0("due_dt")%>"
		.lgStrPrevKey4 = "<%=rs0("splt_seq_no")%>"
				
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hFromPlanDt.value		= "<%=Request("txtFromPlanDt")%>"
		.frm1.hToPlanDt.value		= "<%=Request("txtToPlanDt")%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hProcType.value		= "<%=Request("rdoProcType")%>"
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"

<%		rs0.Close
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
	.DbQueryOk
	
End With	
</Script>	
<%
Set ADF = Nothing
%>
