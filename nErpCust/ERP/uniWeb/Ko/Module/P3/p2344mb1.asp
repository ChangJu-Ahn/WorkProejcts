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
'*  3. Program ID           : p2344mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim strQryMode
Dim i, j

Const C_SHEETMAXROWS = 100

Dim lgStrPrevKey11	' 이전 값 
Dim lgStrPrevKey12

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strItemCd
Dim strTrackingNo
Dim strConvType1
Dim strConvType2
Dim strStartDt
Dim strEndDt
Dim txtPurOrg
Dim cboProdMgr

	lgStrPrevKey11 = Request("lgStrPrevKey11")
	lgStrPrevKey12 = Request("lgStrPrevKey12")
	
	Redim UNISqlId(6)
	Redim UNIValue(6, 11)
	
	UNISqlId(0) = "P2344MB1A"
	UNISqlId(1) = "P2344MB1B"
	UNISqlId(2) = "184000saa"
	UNISqlId(3) = "184000sac"
	UNISqlId(4) = "180000sam"
	UNISqlId(5) = "s0000qa021"
	UNISqlId(6) = "127400saa"
	
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

	IF Request("rdoConvType") = "A" THEN
       strConvType1 = "|"
       strConvType2 = "|"
    ELSEIF Request("rdoConvType") = "NL" then
		strConvType1 = "" & FilterVar("NL", "''", "S") & ""
		strConvType2 = "|"		
    Else	
		strConvType1 = "|"
		strConvType2 = "" & FilterVar("NL", "''", "S") & ""
    END IF

	IF Request("txtStartDt") = "" THEN
	   strStartDt = "|"
	ELSE
	   strStartDt = FilterVar(UniConvDate(Request("txtStartDt"))	, "''", "S")
	END IF

    IF Request("txtEndDt") = "" THEN
    	strEndDt = "|"
    ELSE
    	strEndDt = FilterVar(UniConvDate(Request("txtEndDt"))	, "''", "S")
    END IF   
    
    IF Request("cboProdMgr") = "" THEN
    	cboProdMgr = "|"
    ELSE
    	cboProdMgr = FilterVar(Request("cboProdMgr")	, "''", "S")
    END IF   
    
    IF Request("txtPurOrg") = "" THEN
    	txtPurOrg = "|"
    ELSE
    	txtPurOrg = FilterVar(Trim(Request("txtPurOrg"))	, "''", "S")
    END IF  
			
	UNIValue(0, 0) = "^"	
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 2) = strItemCd	
	UNIValue(0, 3) = strTrackingNo		
	UNIValue(0, 4) = strConvType1
	UNIValue(0, 5) = strConvType2
	UNIValue(0, 6) = "|"
	UNIValue(0, 7) = "|"
	UNIValue(0, 8) = strStartDt
	UNIValue(0, 9) = strEndDt
	UNIValue(0, 10) = cboProdMgr
	IF Request("txtItemGroupCd") = "" Then
		UNIValue(0, 11) = "|"
	Else
		UNIValue(0, 11) = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
		
	UNIValue(1, 0) = "^"
	UNIValue(1, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(1, 2) = strItemCd	
	UNIValue(1, 3) = strTrackingNo		
	UNIValue(1, 4) = strConvType1
	UNIValue(1, 5) = strConvType2
	UNIValue(1, 6) = "|"	
	UNIValue(1, 7) = "|"
	UNIValue(1, 8) = strStartDt
	UNIValue(1, 9) = strEndDt
	UNIValue(1, 10) = txtPurOrg
	IF Request("txtItemGroupCd") = "" Then
		UNIValue(1, 11) = "|"
	Else
		UNIValue(1, 11) = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	UNIValue(2, 0) = FilterVar(Request("txtPlantCd")	, "''", "S")
	UNIValue(3, 0) = FilterVar(Request("txtItemCd")	, "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	UNIValue(5, 0) = FilterVar(Request("txtPurOrg"), "''", "S")
	UNIValue(6, 0) = FilterVar(UCase(Request("txtItemGroupCd")),"''","S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)

	If Not(rs2.EOF AND rs2.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "With parent.frm1" & vbCrLf
				Response.Write ".txtPlantNm.value = """ & ConvSPChars(rs2("plant_nm")) & """" & vbCrLf
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

	If Not(rs3.EOF AND rs3.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs3("item_nm")) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
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
		End If
	End If
      
	If Not(rs6.EOF AND rs6.BOF) Then
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemGroupNm.value = """ & ConvSPChars(rs6("item_group_nm")) & """" & vbCrLf
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
      
	If (rs0.EOF And rs0.BOF) and (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		rs1.Close
		Set rs0 = Nothing
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
		Response.End
	End If
	
%>

<Script Language=vbscript>

Dim LngMaxRow1
Dim LngMaxRow2
Dim strData1
Dim strData2
Dim arrVal1
Dim arrVal2
ReDim arrVal1(0)
ReDim arrVal2(0)

With parent	
	LngMaxRow1 = .frm1.vspdData1.MaxRows
	LngMaxRow2 = .frm1.vspdData2.MaxRows
		
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1 
			If i < C_SHEETMAXROWS THEN
%>
				strData1 = ""
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
				strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PLAN_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("BASIC_UNIT"))%>"			'단위 
				strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("START_PLAN_DT"))%>"
				strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("END_PLAN_DT"))%>"		
<%				IF Trim(rs0("PLAN_STATUS")) ="NL" Then%>
				   strData1 = strData1 & Chr(11) & "Plan"
<%				ELSEIF Trim(rs0("PLAN_STATUS")) = "OP" Then%>
					strData1 = strData1 & Chr(11) & "Open"
<%				ELSEIF Trim(rs0("PLAN_STATUS")) = "RL" Then%>
					strData1 = strData1 & Chr(11) & "Release"
<%				ELSEIF Trim(rs0("PLAN_STATUS")) = "ST" Then%>
					strData1 = strData1 & Chr(11) & "Start"
<%				ELSE%>
					strData1 = strData1 & Chr(11) & "Plan"
<%				END IF%>

				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("ORDER_NO"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("MINOR_NM"))%>"		'생산담당자 
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
				strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
				strData1 = strData1 & Chr(11) & LngMaxRow1 + "<%=i%>"
				strData1 = strData1 & Chr(11) & Chr(12)
				
				ReDim Preserve arrVal1(<%=i%>)
				arrVal1(<%=i%>) = strData1
<%		
				rs0.MoveNext
			END IF	
		Next
%>
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowData Join(arrVal1,"")
		
		.lgStrPrevKey11 = "<%=ConvSPChars(rs0("ITEM_CD"))%>"
		.lgStrPrevKey12 = "<%=ConvSPChars(rs0("PLAN_ORDER_NO"))%>"

<%	End If%>	

		.frm1.hPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.hConvType.value	= "<%=Request("rdoConvType")%>"
		.frm1.hProdMgr.value	= "<%=ConvSPChars(Request("cboProdMgr"))%>"
		.frm1.hPurOrg.value		= "<%=ConvSPChars(Request("txtPurOrg"))%>"
		.frm1.hStartDt.value	= "<%=Trim(Request("txtStartDt"))%>"
		.frm1.hEndDt.value	    = "<%=Trim(Request("txtEndDt"))%>"
		.frm1.hItemGroupCd.value= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		
<%  
	If Not(rs1.EOF And rs1.BOF) Then
		For j=0 to rs1.RecordCount-1 
			IF j < C_SHEETMAXROWS THEN
%>
				strData2 = ""
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("ITEM_CD"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("ITEM_NM"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("SPEC"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("tracking_no"))%>"
				strData2 = strData2 & Chr(11) & "<%=UniConvNumberDBToCompany(rs1("PLAN_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("BASIC_UNIT"))%>"
				strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("START_PLAN_DT"))%>"
				strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("END_PLAN_DT"))%>"

<%				IF Trim(rs1("PLAN_STATUS")) = "N" Then%>
				   strData2 = strData2 & Chr(11) & "Plan"
<%				ELSEIF Trim(rs1("PLAN_STATUS")) = "OP" Then%>
					strData2 = strData2 & Chr(11) & "Open"
<%				ELSEIF Trim(rs1("PLAN_STATUS")) = "RL" Then%>
					strData2 = strData2 & Chr(11) & "Release"
<%				ELSEIF Trim(rs1("PLAN_STATUS")) = "ST" Then%>
					strData2 = strData2 & Chr(11) & "Start"
<%				ELSE%>
					strData2 = strData2 & Chr(11) & "Plan"
<%				END IF%>

				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("ORDER_NO"))%>"
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("PUR_ORG"))%>" 
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_group_cd"))%>"	
				strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_group_nm"))%>"	
				strData2 = strData2 & Chr(11) & LngMaxRow2 + "<%=j%>"
				strData2 = strData2 & Chr(11) & Chr(12)
				
				ReDim Preserve arrVal2(<%=j%>)
				arrVal2(<%=j%>) = strData2
<%		
				rs1.MoveNext
			End IF
		Next
%>
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData Join(arrVal2,"")
		
		.lgStrPrevKey21 = "<%=ConvSPChars(rs1("ITEM_CD"))%>"
		.lgStrPrevKey22 = "<%=ConvSPChars(rs1("PLAN_ORDER_NO"))%>"
		
<%	End If%>

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
