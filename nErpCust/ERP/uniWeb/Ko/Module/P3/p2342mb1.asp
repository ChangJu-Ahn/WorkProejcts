<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call LoadBasisGlobalInf
   Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB") 
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2342mb1.asp
'*  4. Program Name         : MRP Base
'*  5. Program Desc         : query MRP Base1
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim strQryMode

Const C_SHEETMAXROWS_D = 100

Dim lgStrPrevKey11	' 이전 값 
Dim lgStrPrevKey12	' 이전 값 
Dim i

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strItemCd
Dim strTrackingNo
Dim strProcType
Dim cboMrpMgr
Dim cboProdMgr
Dim txtPurOrg
Dim txtPurGrp
Dim txtSuppl
	
	
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write "	parent.frm1.txtItemNm.value = """"" & vbCrLf
	Response.Write "	parent.frm1.txtPurOrgNm.value = """"" & vbCrLf
	Response.Write "	parent.frm1.txtPurGrpNm.value = """"" & vbCrLf
	Response.Write "	parent.frm1.txtSupplNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	
	lgStrPrevKey11 = UCase(Trim(Request("lgStrPrevKey11")))
	lgStrPrevKey12 = UCase(Trim(Request("lgStrPrevKey12")))
	
	Redim UNISqlId(6)
	Redim UNIValue(6, 8)
	
	UNISqlId(0) = "185200saa"
	UNISqlId(1) = "184000saa"
	UNISqlId(2) = "184000sac"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "s0000qa021"			'구매조직 
	UNISqlId(5) = "M3111QA104"			'구매그룹 
	UNISqlId(6) = "s0000qa024"			'공급처 
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd =FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	END IF
	
	IF Request("txtTrackingNo") = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(Trim(Request("txtTrackingNo"))	, "''", "S")
	END IF
	
	IF Request("rdoProcType") = "A" THEN
       strProcType = "|"
    ELSE
       strProcType = FilterVar(Request("rdoProcType")	, "''", "S")
    END IF

	IF Request("cboMrpMgr") = "" Then
	   cboMrpMgr = "|"
	ELSE
	   cboMrpMgr = FilterVar(Request("cboMrpMgr")	, "''", "S")
	END IF	
	
	IF Request("cboProdMgr") = "" THEN
    	cboProdMgr = "|"
    ELSE
    	cboProdMgr = FilterVar(Request("cboProdMgr")	, "''", "S")
    END IF   
    
    If Request("txtPurOrg") = "" Then
    	txtPurOrg = "|"
    Else
    	txtPurOrg = FilterVar(Request("txtPurOrg"), "''", "S")
    End If  

    If Request("txtPurGrp") = "" Then
    	txtPurGrp = "|"
    Else
    	txtPurGrp = FilterVar(Request("txtPurGrp"), "''", "S")
    End If  

    If Request("txtSuppl") = "" Then
    	txtSuppl = "|"
    Else
    	txtSuppl = FilterVar(Request("txtSuppl"), "''", "S")
    End If  
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")

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
		If strTrackingNo = "|" Then
			UNIValue(0, 2) = "((a.item_cd = " & FilterVar(Trim(lgStrPrevKey11)	, "''", "S") & _
							" and a.tracking_no >= " & FilterVar(Trim(lgStrPrevKey12)	, "''", "S") & _
							") or a.item_cd > " & FilterVar(Trim(lgStrPrevKey11)	, "''", "S") & ")"
		Else
			UNIValue(0, 2) = "(a.item_cd >= " & FilterVar(Trim(lgStrPrevKey11)	, "''", "S") & _
							" and a.tracking_no = " & strTrackingNo & ")"
		End If
	End Select
	
	UNIValue(0, 3) = strProcType
	UNIValue(0, 4) = cboMrpMgr
    UNIValue(0, 5) = cboProdMgr
    UNIValue(0, 6) = txtPurOrg
    UNIValue(0, 7) = txtPurGrp
    UNIValue(0, 8) = txtSuppl
	
	UNIValue(1, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(2, 0) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(3, 0) = FilterVar(Request("txtTrackingNo"), "''", "S")
	UNIValue(4, 0) = FilterVar(Request("txtPurOrg"),"''","S")
	UNIValue(5, 0) = FilterVar(Request("txtPurGrp"), "''", "S")
	UNIValue(6, 0) = FilterVar(Request("txtSuppl"),"''","S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1" 
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)

	If Not(rs1.EOF AND rs1.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "	parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("plant_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "	parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If

	If Not(rs2.EOF AND rs2.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "	parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("item_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
	
	IF Request("txtTrackingNo") <> "" Then
		If (rs3.EOF AND rs3.BOF) Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "	parent.frm1.txtTrackingNo.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			
			Response.End
		End If
	End If
	
	IF Request("txtPurOrg") <> "" Then
		If (rs4.EOF AND rs4.BOF) Then
			Call DisplayMsgBox("125200", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurOrg.focus" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurOrgNm.value = """ & ConvSPChars(rs4("PUR_ORG_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			
		End If
	End If

	IF Request("txtPurGrp") <> "" Then
		If (rs5.EOF AND rs5.BOF) Then
			Call DisplayMsgBox("125100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurGrp.focus" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurGrpNm.value = """ & ConvSPChars(rs5("PUR_GRP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	End If

	IF Request("txtSuppl") <> "" Then
		If (rs6.EOF AND rs6.BOF) Then
			Call DisplayMsgBox("229927", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSuppl.focus" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSupplNm.value = """ & ConvSPChars(rs6("BP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf	
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
		rs5.Close
		Set rs5 = Nothing				
		rs6.Close
		Set rs6 = Nothing						
		Response.End
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
	If Not(rs0.EOF And rs0.BOF) Then
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"			
				strData = strData & Chr(11) & "<%=rs0("tracking_flg")%>"
<%			IF Trim(rs0("PROCUR_TYPE")) = "M" Then
%>
			   strData = strData & Chr(11) & "제조"
<%			ELSEIF Trim(rs0("PROCUR_TYPE")) = "P" Then
%>
			   strData = strData & Chr(11) & "구매"
<%			ELSE
%>
			   strData = strData & Chr(11) & "외주"
<%			END IF
%>
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MRP_MGR"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PROD_MGR"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_GRP"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_GRP_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
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
		
		.lgStrPrevKey11	= "<%=ConvSPChars(rs0("item_cd"))%>"
		.lgStrPrevKey12 = "<%=ConvSPChars(rs0("tracking_no"))%>"
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value = "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"	
	.frm1.hProcType.value   = "<%=Request("rdoProcType")%>"
	.frm1.hMrpMgr.value = "<%=ConvSPChars(Request("cboMrpMgr"))%>"
	.frm1.hProdMgr.value = "<%=ConvSPChars(Request("cboProdMgr"))%>"
	.frm1.hPurOrg.value = "<%=ConvSPChars(Request("txtPurOrg"))%>"
	.frm1.hPurGrp.value = "<%=ConvSPChars(Request("txtPurGrp"))%>"
	.frm1.hSuppl.value = "<%=ConvSPChars(Request("txtSupple"))%>"	

	.DbQueryOk
End With	
</Script>	
<%
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
	rs5.Close
	Set rs5 = Nothing				
	rs6.Close
	Set rs6 = Nothing	
	
	Set ADF = Nothing
%>
