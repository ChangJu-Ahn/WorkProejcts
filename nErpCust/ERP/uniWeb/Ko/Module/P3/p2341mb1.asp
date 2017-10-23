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
'*  3. Program ID           : p2341mb1.asp
'*  4. Program Name         : MRP Results
'*  5. Program Desc         : query MRP Results
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2003-12-06
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7
Dim strQryMode

Const C_SHEETMAXROWS = 100

Dim lgStrPrevKey	' ���� �� 
Dim lgStrPrevKey2	' ���� �� (Tracking_No)
Dim lgStrPrevKey3	' ���� �� (Due_Dt)
Dim lgStrPrevKey4	' ���� �� (Split_Seq_No)
Dim lgStrPrevKey5	' ���� �� (MPS_No)
Dim i

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strItemCd
Dim strTrackingNo
Dim strProcType
Dim strFromPlanDt
Dim strToPlanDt
Dim cboMrpMgr
Dim cboProdMgr
Dim txtPurOrg
Dim txtPurGrp
Dim txtSuppl
	
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtItemGroupNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtPurOrgNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtPurGrpNm.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtSupplNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf

	Redim UNISqlId(7)
	Redim UNIValue(7, 11)
	
	UNISqlId(0) = "P2341MB1"
	UNISqlId(1) = "184000saa"
	UNISqlId(2) = "184000sac"
	UNISqlId(3) = "180000sam"
	UNISqlId(4) = "127400saa"
	UNISqlId(5) = "s0000qa021"			'�������� 
	UNISqlId(6) = "M3111QA104"			'���ű׷� 
	UNISqlId(7) = "s0000qa024"			'����ó 
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	END IF
	
	IF Request("txtTrackingNo") = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(Request("txtTrackingNo")	, "''", "S")
	END IF

    IF Request("rdoProcType") = "A" THEN
       strProcType = "|"
    ELSE
       strProcType = FilterVar(Request("rdoProcType")	, "''", "S")
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
	UNIValue(0, 2) = strItemCd
	
	Select Case strQryMode	
	Case CStr(OPMD_CMODE)
		If strItemCd = "|" and strTrackingNo = "|" Then             ' ǰ��, Tracking No �Ѵ� ���� 
			UNIValue(0, 2) = "|"
		ElseIf strItemCd <> "|" and strTrackingNo = "|" Then        ' ǰ�� �Է� 
			UNIValue(0, 2) = "a.item_cd >= " & strItemCd
		ElseIf strItemCd = "|" and strTrackingNo <> "|" Then        ' Tracking No�� �Է� 
			UNIValue(0, 2) = "a.tracking_no = " & strTrackingNo
		Else                                                        ' ǰ��, Tracking No �Ѵ� �Է� 
			UNIValue(0, 2) = "(a.item_cd >= " & strItemCd & " and a.tracking_no = " & strTrackingNo & ")"
		End If			
	Case CStr(OPMD_UMODE)
		lgStrPrevKey = FilterVar(Request("lgStrPrevKey")	, "''", "S")
		lgStrPrevKey2 = FilterVar(Request("lgStrPrevKey2")	, "''", "S")
		lgStrPrevKey3 = FilterVar(Request("lgStrPrevKey3")	, "''", "S")
		lgStrPrevKey4 = FilterVar(Request("lgStrPrevKey4")	, "''", "S")
		lgStrPrevKey5 = FilterVar(Request("lgStrPrevKey5")	, "''", "S")


		If strTrackingNo = "|" Then
			UNIValue(0, 2) = "((a.item_cd = " & lgStrPrevKey & " and a.due_dt = " & lgStrPrevKey3 & _
							" and splt_seq_no >= " & lgStrPrevKey4 & " and a.tracking_no >= " & lgStrPrevKey2 & _
							" and mps_no >= " & lgStrPrevKey5 & _
			                 ") or (a.item_cd >= " & lgStrPrevKey & " and a.due_dt = " & lgStrPrevKey3 & _
			                 " and splt_seq_no > " & lgStrPrevKey4 & _ 
			                 ") or (a.item_cd >= " & lgStrPrevKey & " and a.due_dt > " & lgStrPrevKey3 & _
			                 ") or a.item_cd > " & lgStrPrevKey & ")"
		Else
			UNIValue(0, 2) = "((a.item_cd = " & lgStrPrevKey & " and a.tracking_no = " & strTrackingNo & _
							" and a.due_dt = " & lgStrPrevKey3 & " and splt_seq_no >= " & lgStrPrevKey4 & _
							" and a.mps_no >= " & lgStrPrevKey5 & _
			                 ") or (a.item_cd >= " & lgStrPrevKey & " and a.tracking_no = " & strTrackingNo & _
			                 " and a.due_dt > " & lgStrPrevKey3 & _
			                 ") or (a.item_cd > " & lgStrPrevKey & " and a.tracking_no = " & strTrackingNo & "))"		
		End If	
	End Select
		
	UNIValue(0, 3) = strProcType
	UNIValue(0, 4) = strFromPlanDt
	UNIValue(0, 5) = strToPlanDt
	IF Request("txtItemGroupCd") = "" Then
		UNIValue(0,6) = "|"
	Else
		UNIValue(0,6) = "c.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	UNIValue(0, 7) = cboMrpMgr
    UNIValue(0, 8) = cboProdMgr
    UNIValue(0, 9) = txtPurOrg
    UNIValue(0, 10) = txtPurGrp
    UNIValue(0, 11) = txtSuppl
    
	'---------------------------------------------------------------------------------------------------------
	
	UNIValue(1, 0) = FilterVar(Request("txtPlantCd")	, "''", "S")
	UNIValue(2, 0) = FilterVar(Request("txtItemCd")	, "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtItemGroupCd")),"''","S")
	UNIValue(5, 0) = FilterVar(Request("txtPurOrg"),"''","S")
	UNIValue(6, 0) = FilterVar(Request("txtPurGrp"), "''", "S")
	UNIValue(7, 0) = FilterVar(Request("txtSuppl"),"''","S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7)

	If Not(rs1.EOF AND rs1.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("plant_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		
		Response.End 
	End If

	If Not(rs2.EOF AND rs2.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs2("item_nm")) & """" & vbCrLf
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
		IF Request("txtItemGroupCd") <> "" Then
			Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemGroupCd.focus" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.End 
		End If
	End If
	
	IF Request("txtPurOrg") <> "" Then
		If (rs5.EOF AND rs5.BOF) Then
			Call DisplayMsgBox("125200", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurOrg.focus" & vbCrLf	'��: ȭ�� ó�� ASP �� ��Ī�� 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurOrgNm.value = """ & ConvSPChars(rs5("PUR_ORG_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			
		End If
	End If

	IF Request("txtPurGrp") <> "" Then
		If (rs6.EOF AND rs6.BOF) Then
			Call DisplayMsgBox("125100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurGrp.focus" & vbCrLf	'��: ȭ�� ó�� ASP �� ��Ī�� 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPurGrpNm.value = """ & ConvSPChars(rs6("PUR_GRP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	End If

	IF Request("txtSuppl") <> "" Then
		If (rs7.EOF AND rs7.BOF) Then
			Call DisplayMsgBox("229927", vbOKOnly, "", "", I_MKSCRIPT)
			
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSuppl.focus" & vbCrLf	'��: ȭ�� ó�� ASP �� ��Ī�� 
			Response.Write "</Script>" & vbCrLf
		
			Response.End
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSupplNm.value = """ & ConvSPChars(rs7("BP_NM")) & """" & vbCrLf
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
		rs7.Close
		Set rs7 = Nothing	
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
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("start_dt"))%>"
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("due_dt"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("basic_unit"))%>"

<%			IF Trim(rs0("procur_type")) = "M" Then%>
			   strData = strData & Chr(11) & "����"
<%			ELSEIF Trim(rs0("procur_type")) = "P" Then%>
			   strData = strData & Chr(11) & "����"
<%			ELSE%>
			   strData = strData & Chr(11) & "����"
<%			END IF%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MRP_MGR"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PROD_MGR"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_GRP"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_GRP_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
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
		.lgStrPrevKey5 = "<%=rs0("mps_no")%>"
		
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hTrackingNo.value		= "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hProcType.value		= "<%=Request("rdoProcType")%>"
		.frm1.hFromPlanDt.value		= "<%=Request("txtFromPlanDt")%>"
		.frm1.hToPlanDt.value		= "<%=Request("txtToPlanDt")%>"
		.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		.frm1.hMrpMgr.value			= "<%=ConvSPChars(Request("cboMrpMgr"))%>"
		.frm1.hProdMgr.value		= "<%=ConvSPChars(Request("cboProdMgr"))%>"
		.frm1.hPurOrg.value			= "<%=ConvSPChars(Request("txtPurOrg"))%>"
		.frm1.hPurGrp.value			= "<%=ConvSPChars(Request("txtPurGrp"))%>"
		.frm1.hSuppl.value			= "<%=ConvSPChars(Request("txtSupple"))%>"
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
		rs7.Close
		Set rs7 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing
%>
