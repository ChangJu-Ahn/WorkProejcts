<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4413rb1.asp
'*  4. Program Name			: List Rework Order History (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2003/02/20
'*  8. Modified date(Last)	: 2003/02/22
'*  9. Modifier (First)		: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4					'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim i
Dim strFlag
Dim strProdtOrderNo 
Dim strOprNo
Dim strOrdOprFlag

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================

Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

On Error Resume Next
Err.Clear
																	'☜: Protect system from crashing
	'// GET Plant Name & Item Name	
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)
		
	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs3, rs4)
    
	' Plant 명 Display      
	If (rs3.EOF And rs3.BOF) Then
		rs3.Close
		Set rs3 = Nothing
		strFlag = "ERROR_PLANT"	
	Else
		%>
		<Script Language=vbscript>
			parent.txtPlantNm.value = "<%=ConvSPChars(rs3("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs3.Close
		Set rs3 = Nothing
	End If
	
	'Item Name Display
	If (rs4.EOF And rs4.BOF) Then
		rs4.Close
		Set rs4 = Nothing
		strFlag = "ERROR_ITEM"
	Else
		%>
		<Script Language=vbscript>
			parent.txtItemNm.value = "<%=ConvSPChars(rs4("ITEM_NM"))%>"
		</Script>	
		<%
		rs4.Close
		Set rs4 = Nothing
	End If
	
	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			Response.End
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.End
		End If
	End IF
	
	Set ADF = Nothing
	
	'// GET REWORK_ORDER_NO
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
		
	UNISqlId(0) = "P4413RB1P"
	UNIValue(0, 0) = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2)
    
    If Trim(Request("txtOprNo")) = "" Then
		strOrdOprFlag = "ORD"
    Else
		strOrdOprFlag = "OPR"
    End If
    
	If (rs2.EOF And rs2.BOF) Then
		strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
		Select Case strOrdOprFlag
			Case "ORD"
				strOprNo = "|"
			Case "OPR"
				strOprNo = UCase(Trim(Request("txtOprNo")))
		End Select 
		
		rs2.Close
		Set rs2 = Nothing	
	Else
		strProdtOrderNo = FilterVar(UCase(rs2("ORIGINAL_ORDER_NO")), "''", "S")
		Select Case strOrdOprFlag
			Case "ORD"
				strOprNo = "|"
			Case "OPR"
				strOprNo = FilterVar(UCase(rs2("ORIGINAL_OPR_NO")), "''", "S")
		End Select 
		
		rs2.Close
		Set rs2 = Nothing
	End If
	
	Set ADF = Nothing
	
	'// QUERY PROD ORDER HEADER
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "P4413RB1Q"
	
	UNIValue(0, 0) = strProdtOrderNo
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
	
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
	'// QUERY REWORK ORDER HISTORY
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 2)
	
	UNISqlId(0) = "P4413RB1H"
	UNISqlId(1) = "P4413RB1D"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strProdtOrderNo
	UNIValue(0, 2) = strOprNo
	UNIValue(1, 0) = "^"
	UNIValue(1, 1) = strProdtOrderNo
	UNIValue(1, 2) = strOprNo
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2, rs0)
    
    If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("189239", vbOKOnly, "", "", I_MKSCRIPT)
		rs2.Close
		Set rs0 = Nothing
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	  
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("189239", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set rs2 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
%>
<Script Language=vbscript>
	Dim TmpBuffer1
    Dim TmpBuffer2
    Dim iTotalStr
    Dim LngMaxRow
    Dim strData
	
    With parent												'☜: 화면 처리 ASP 를 지칭함 
		
	 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow
		
<%  
		If Not(rs0.EOF And rs0.BOF) Then
%>	
		
			.txtOriginalOrderNo.value = "<%=ConvSPChars(rs1("PRODT_ORDER_NO"))%>"
			
			.txtOrderQty.value = "<%=UniConvNumberDBToCompany(rs1("PRODT_ORDER_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			.txtProdQty.text = "<%=UniNumClientFormat(rs1("PROD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
			.txtOrderUnit.value = "<%=ConvSPChars(rs1("PRODT_ORDER_UNIT"))%>"
			.txtDefectQty.text = "<%=UniNumClientFormat(rs1("BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
			.txtInspDefectQty.text = "<%=UniNumClientFormat(rs1("INSP_BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
			.txtPlanStratDt.text = "<%=UNIDateClientFormat(rs1("PLAN_START_DT"))%>"
			.txtPlanEndDt.text = "<%=UNIDateClientFormat(rs1("PLAN_COMPT_DT"))%>"
			.txtTrackingNo.value =  "<%=ConvSPChars(rs1("TRACKING_NO"))%>"
			.txtOrderStatus.value = "<%=rs1("STATUS")%>"
			
			Redim TmpBuffer1(<%=rs2.RecordCount-1%>)
<%		
			For i=0 to rs2.RecordCount-1
%>
				strData = ""
				
				If UCase(Trim(<%=rs2("RE_WORK_FLG")%>)) = "Y" Then
					strData = strData & Chr(11) & "재작업"
				Else 
					strData = strData & Chr(11) & "작업"
				End If
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("PRODT_ORDER_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("PARENT_ORDER_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("PARENT_OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=rs2("STATUS")%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("PRODT_ORDER_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("PROD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("REWORK_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("PRODT_ORDER_UNIT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs2("PLAN_START_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs2("PLAN_COMPT_DT"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs2("INSP_BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs2("TRACKING_NO"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer1(<%=i%>) = strData
<%		
				rs2.MoveNext
				
			Next
%>
			
			Redim TmpBuffer2(<%=rs0.RecordCount-1%>)
<%		
			For i=0 to rs0.RecordCount-1
%>
				strData = ""
				
				If UCase(Trim("<%=rs0("RE_WORK_FLG")%>")) = "Y" Then
					strData = strData & Chr(11) & "재작업"
				Else 
					strData = strData & Chr(11) & "작업"
				End If
				
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PARENT_ORDER_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PARENT_OPR_NO"))%>"
				strData = strData & Chr(11) & "<%=rs0("STATUS")%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODT_ORDER_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PROD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REWORK_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_START_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_COMPT_DT"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("INSP_BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i + rs2.RecordCount %>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer2(<%=i%>) = strData
<%		
				rs0.MoveNext
				
			Next
%>
		iTotalStr = Join(TmpBuffer1,"") & Join(TmpBuffer2, "")
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr

<%	
		End If
		
		rs0.Close
		rs1.close
		rs2.close
		Set rs0 = Nothing
		Set rs1 = Nothing
		Set rs2 = Nothing
%>	
		
		.DbQueryOk(LngMaxRow)
		
    End With
</Script>	
<%    
    Set ADF = Nothing
%>
