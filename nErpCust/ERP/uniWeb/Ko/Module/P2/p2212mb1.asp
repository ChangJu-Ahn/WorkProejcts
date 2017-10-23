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
'*  3. Program ID           : p2212mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3
Dim strQryMode
Dim i

Const C_SHEETMAXROWS = 100

Dim lgStrPrevKey11	' 이전 값 
Dim lgStrPrevKey12	' 이전 값 

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strPlantCd, strItemCd, strTrackingNo

	lgStrPrevKey11 = FilterVar(Request("lgStrPrevKey11"), "''", "S")
	lgStrPrevKey12 = FilterVar(Request("lgStrPrevKey12"), "''", "S")
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	
	Redim UNISqlId(3)
	Redim UNIValue(3, 2)
	
	UNISqlId(0) = "184200saa"
	UNISqlId(1) = "184000saa"	
	UNISqlId(2) = "184000sac"
	UNISqlId(3) = "180000sam"
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   StrItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	END IF
	
	IF Request("txtTrackingNo") = "" Then
	   strTrackingNo = "|"
	ELSE
	   StrTrackingNo = FilterVar(Trim(Request("txtTrackingNo"))	, "''", "S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	
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
		Case Cstr(OPMD_UMODE)
			If strTrackingNo = "|" Then 
				UNIValue(0, 2) = "((a.item_cd = " & lgStrPrevKey11 & " and a.tracking_no >= " & lgStrPrevKey12 & _
								") or a.item_cd > " & lgStrPrevKey11 & ")"
			Else
				UNIValue(0, 2) = "(a.item_cd >= " & lgStrPrevKey11 & " and a.tracking_no = " & strTrackingNo & ")"
			End If
	End Select
	
	UNIValue(1, 0) = strPlantCd
	UNIValue(2, 0) = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
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
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
		
<%  
    For i=0 to rs0.RecordCount-1 
		If i < C_SHEETMAXROWS Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			ReDim Preserve arrVal(<%=i%>)
			arrVal(<%=i%>) = strData
<%		
			rs0.MoveNext
		End If
	Next
%>
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowData Join(arrVal,"")
		
		.lgStrPrevKey11 = "<%=ConvSPChars(rs0("item_cd"))%>"
		.lgStrPrevKey12 = "<%=ConvSPChars(rs0("tracking_no"))%>"		

		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
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
