<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : Preliminary Delivery Order Status
'*  3. Program ID           : mc900qb1
'*  4. Program Name         : �������ô����ȸ 
'*  5. Program Desc         : List Preliminary Delivery Order Status
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/03/05
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "M", "NOCOOKIE","QB")
On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter ���� 
Dim strQryMode								'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Const C_SHEETMAXROWS = 50

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strPlantCd
Dim strPoFrDt
Dim strPoToDt
Dim strItemCd
Dim strBpCd
Dim strFlag
Dim strPoNo
Dim strPoSeqNo
Dim strVal
Dim strVal2
Dim PvArr

Err.Clear																	'��: Protect system from crashing

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)

	UNISqlId(0) = "mc900qb1"
	
	
		
	If Request("txtPlantCd") = "" Then
		strPlantCd = "|"
	Else 
		StrPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	End If

	IF Request("txtPoFrDt") = "" Then
		strPoFrDt = "|"
	Else
		strPoFrDt = " " & FilterVar(UNIConvDate(Request("txtPoFrDt")), "''", "S") & ""
	End IF

	IF Request("txtPoToDt") = "" Then
		strPoToDt = "|"
	Else
		strPoToDt = " " & FilterVar(UNIConvDate(Request("txtPoToDt")), "''", "S") & ""
	End IF

	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtBpCd") = "" Then
		strBpCd = "|"
	Else
		strBpCd = FilterVar(UCase(Request("txtBpCd")), "''", "S")
	End IF

	If Request("cboDlvyOrderFlag") = "C" Then
		strVal = strVal & " A.BASE_QTY = A.BASE_DLY_QTY "
	Elseif Request("cboDlvyOrderFlag") = "I" Then
		strVal = strVal & " A.BASE_QTY <> A.BASE_DLY_QTY AND A.PO_QTY > A.BASE_RCPT_QTY"
	Elseif Request("cboDlvyOrderFlag") = "F" Then
		strVal = strVal & " A.PO_QTY = A.BASE_RCPT_QTY"
	Else
		strVal = "|"
	End if
	
	If Request("lgStrPrevKey1") = "" Then
		strPoNo = "|"
	Else
		strPoNo = FilterVar(UCase(Request("lgStrPrevKey1")), "''", "S")
	End If	
	
	If Request("lgStrPrevKey2") = "" Then
		strPoSeqNo = "|"
	Else
		strPoSeqNo = FilterVar(UCase(Request("lgStrPrevKey2")), "''", "S")
	End If
	
	If strPoNo <> "|" and strPoSeqNo <> "|" Then
		strVal2 = strVal2 & "((A.po_no > " & strPoNo & ") OR (A.po_no = " & strPoNo & " and  A.po_seq_no >= " & strPoSeqNo & "))"
	Else 
		strVal2 = "|"
	End if
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strVal 
	UNIValue(0, 2) = strPlantCd
	UNIValue(0, 3) = strItemCd
	UNIValue(0, 4) = strBpCd 
	UNIValue(0, 5) = strPoFrDt 
	UNIValue(0, 6) = strPoToDt
	UNIValue(0, 7) = strVal2
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.frm1.txtPlantCd.focus " & vbCr
		Response.Write "	Set Parent.gActiveElement = Parent.document.activeElement    " & vbCr
		Response.Write "</Script>" & vbCr
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow

<%  
    ReDim PvArr(C_SHEETMAXROWS - 1)
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS Then 
%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Plant_Cd"))%>"						
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Plant_Nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Po_No"))%>"
			strData = strData & Chr(11) & "<%=rs0("Po_Seq_No")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Cd"))%>"							
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Item_Nm"))%>"							
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"								
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Dlvy_Dt"))%>"						
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Po_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Po_Unit"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Base_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Base_Unit"))%>"							
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Po_Dly_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Po_Rcpt_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Base_Dly_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("Base_Rcpt_Qty"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Sl_Cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Bp_Cd"))%>"								
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Bp_Nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Pur_Org"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Pur_Grp"))%>"								
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("Po_Dt"))%>"					
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Pr_No"))%>"				
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Tracking_No"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Procure_Type"))%>"						
			
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)

<%		

            PvArr(i) = strData	
			strData = ""

			rs0.MoveNext
			End If
		Next
		strData  = Join(PvArr, "")
%>
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData
		
		.lgStrPrevKey1 = "<%=Trim(rs0("PO_NO"))%>"
		.lgStrPrevKey2 = "<%=Trim(rs0("Po_Seq_No"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0) and .lgStrPrevKey1 <> "" and .lgStrPrevKey2 <> "" Then	
		Call .InitData(LngMaxRow)
		.DbQuery
	Else
		.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hPoFrDt.value			= "<%=UNIDateClientFormat(Request("txtPoFrDt"))%>"
		.frm1.hPoToDt.value			= "<%=UNIDateClientFormat(Request("txtPoToDt"))%>"
		.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hBpCd.value			= "<%=ConvSPChars(Request("txtBpCd"))%>"
		
		.DbQueryOk(LngMaxRow+1)
	End If

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
