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
'*  3. Program ID           : p2346mb1.asp
'*  4. Program Name         : MRP Partial Conversion
'*  5. Program Desc         : query MRP
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")  

On Error Resume Next

Dim ADF	
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8
Dim strQryMode
Dim i

Const C_SHEETMAXROWS = 100

Dim lgStrPrevKey11	' 이전 값 
Dim lgStrPrevKey12

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strItemCd
Dim strTrackingNo
Dim strProcType
Dim cboMrpMgr
Dim txtPurOrg
Dim cboProdMgr
Dim txtPurGrp
Dim txtSuppl

	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write		"parent.frm1.txtPlantNm.value = """"" & vbCrLf
	Response.Write		"parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
	Response.Write		"parent.frm1.txtPurOrgNm.value = """"" & vbCrLf
	Response.Write		"parent.frm1.txtPurGrpNm.value = """"" & vbCrLf
	Response.Write		"parent.frm1.txtSupplNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	lgStrPrevKey11 = UCase(Trim(Request("lgStrPrevKey1")))
	lgStrPrevKey12 = UCase(Trim(Request("lgStrPrevKey2")))

	Redim UNISqlId(8)
	Redim UNIValue(8, 15)
	
	UNISqlId(0) = "P2346MB1"
	UNISqlId(1) = "185000saa"
	UNISqlId(2) = "184000saa"
	UNISqlId(3) = "184000sac"
	UNISqlId(4) = "180000sam"
	UNISqlId(5) = "s0000qa021"
	UNISqlId(6) = "127400saa"
	UNISqlId(7) = "M3111QA104"			'구매그룹 
	UNISqlId(8) = "s0000qa024"			'공급처 
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
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
			UNIValue(0, 2) = strItemCd
		Case CStr(OPMD_UMODE)
			UNIValue(0, 2) = FilterVar(lgStrPrevKey11, "''", "S")

	End Select

	UNIValue(0, 3) = strTrackingNo
	UNIValue(0, 4) = strProcType
	UNIValue(0, 5) = cboMrpMgr
	
	IF Request("txtStartDt") = "" THEN
	   UNIValue(0, 6) = "|"
	ELSE
	   UNIValue(0, 6) = FilterVar(UniConvDate(Request("txtStartDt"))	, "''", "S")
	END IF
	
	IF Request("txtEndDt") = "" THEN
	   UNIValue(0, 7) = "|"
	ELSE
	   UNIValue(0, 7) = FilterVar(UniConvDate(Request("txtEndDt"))	, "''", "S")
	END IF
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 8) = "|"
			UNIValue(0, 9) = "|"
		Case CStr(OPMD_UMODE)
			UNIValue(0, 8) = "a.item_cd > " & FilterVar(lgStrPrevKey11	, "''", "S") & " or (a.item_cd = " & FilterVar(lgStrPrevKey11	, "''", "S")	
			UNIValue(0, 9) = FilterVar(Trim(lgStrPrevKey12)	, "''", "S")
	End Select
	
	UNIValue(0, 10) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
    UNIValue(0, 11) = cboProdMgr
    UNIValue(0, 12) = txtPurOrg

	IF Request("txtItemGroupCd") = "" Then
		UNIValue(0, 13) = "|"
	Else
		UNIValue(0, 13) = "c.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
    UNIValue(0, 14) = txtPurGrp
    UNIValue(0, 15) = txtSuppl
    
    '=================================================================================================================
    
	UNIValue(1, 0) = FilterVar(Request("txtPlantCd")	, "''", "S")
	UNIValue(1, 1) = FilterVar(Request("txtPlantCd")	, "''", "S")
	UNIValue(2, 0) = FilterVar(Request("txtPlantCd")	, "''", "S")
	UNIValue(3, 0) = FilterVar(Request("txtItemCd")	, "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(5, 0) = FilterVar(Request("txtPurOrg"),"''","S")
	UNIValue(6, 0) = FilterVar(UCase(Request("txtItemGroupCd")),"''","S")
	UNIValue(7, 0) = FilterVar(Request("txtPurGrp"), "''", "S")
	UNIValue(8, 0) = FilterVar(Request("txtSuppl"),"''","S")
	
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8)
      
	If Not(rs2.EOF AND rs2.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs2("plant_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If


	If Not(rs3.EOF AND rs3.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs3("item_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
	
	IF Request("txtTrackingNo") <> "" Then
		If (rs4.EOF AND rs4.BOF) Then
			Call DisplayMsgBox("203045", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtTrackingNo.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		End If
	End If
	
	IF Request("txtPurOrg") <> "" Then
		If (rs5.EOF AND rs5.BOF) Then
			Call DisplayMsgBox("125200", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurOrg.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurOrgNm.value = """ & ConvSPChars(rs5("PUR_ORG_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf	
		End If
	End If

	If Not(rs6.EOF AND rs6.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs6("item_group_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		IF Request("txtItemGroupCd") <> "" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End 
		End If
	End If
	
	IF Request("txtPurGrp") <> "" Then
		If (rs7.EOF AND rs7.BOF) Then
			Call DisplayMsgBox("125100", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurGrp.focus" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurGrpNm.value = """ & ConvSPChars(rs7("PUR_GRP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf	
		End If
	End If

	IF Request("txtSuppl") <> "" Then
		If (rs8.EOF AND rs8.BOF) Then
			Call DisplayMsgBox("229927", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSuppl.focus" & vbCrLf	'☜: 화면 처리 ASP 를 지칭함 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSupplNm.value = """ & ConvSPChars(rs8("BP_NM")) & """" & vbCrLf
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
		Set rs1 = Nothing
		rs3.Close
		Set rs3 = Nothing	
		rs4.Close
		Set rs4 = Nothing			
		rs5.Close
		Set rs5 = Nothing
		rs6.Close
		Set rs6 = Nothing
		rs7.Close
		Set rs7 = Nothing
		rs8.Close
		Set rs8 = Nothing
		Response.End
		Set ADF = Nothing
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
		IF i < C_SHEETMAXROWS Then
%>
			strData = ""
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("start_plan_dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("end_plan_dt"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"
			
<%			IF Trim(rs0("procur_type")) = "M" Then
%>
				strData = strData & Chr(11) & "제조"
<%			ELSEIF Trim(rs0("procur_type")) = "P" Then
%>
			    strData = strData & Chr(11) & "구매"
<%			ELSE
%>
				strData = strData & Chr(11) & "외주"
<%			END IF
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("plan_order_no"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("mrp_mgr"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prod_mgr"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_GRP"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_GRP_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
			strData = strData & Chr(11) & ""	
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
		.ggoSpread.SSShowDataByClip Join(arrVal,"")
		
		.lgStrPrevKey1 = "<%=ConvSPChars(rs0("item_cd"))%>"
		.lgStrPrevKey2 = "<%=ConvSPChars(rs0("plan_order_no"))%>"

        .frm1.hmrpno.value			= "<%=ConvSPChars(rs1("run_no"))%>"        
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hStartDt.value		= "<%=Request("txtStartDt")%>"
		.frm1.hEndDt.value			= "<%=Request("txtEndDt")%>"			
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hProdMgr.value		= "<%=ConvSPChars(Request("cboProdMgr"))%>"
		.frm1.hMrpMgr.value			= "<%=ConvSPChars(Request("cboMrpMgr"))%>"
		.frm1.hPurOrg.value			= "<%=ConvSPChars(Request("txtPurOrg"))%>"
		.frm1.hPurGrp.value			= "<%=ConvSPChars(Request("txtPurGrp"))%>"
		.frm1.hSuppl.value			= "<%=ConvSPChars(Request("txtSupple"))%>"
		.frm1.hProcType.value		= "<%=Request("rdoProcType")%>"
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
<%			
		rs0.Close
		Set rs0 = Nothing

		rs1.Close
		Set rs1 = Nothing	
		rs2.Close
		Set rs1 = Nothing
		rs3.Close
		Set rs3 = Nothing
		rs4.Close
		Set rs4 = Nothing	
		rs5.Close
		Set rs5 = Nothing
		rs6.Close
		Set rs6 = Nothing
		rs7.Close
		Set rs6 = Nothing
		rs8.Close
		Set rs6 = Nothing
%>
	.DbQueryOk(LngMaxRow + 1)
End With	
</Script>	
<%
Set ADF = Nothing
%>
