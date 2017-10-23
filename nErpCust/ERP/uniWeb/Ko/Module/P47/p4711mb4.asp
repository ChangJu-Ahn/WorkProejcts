<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711mb4.asp
'*  4. Program Name         :
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001/12/01
'*  7. Modified date(Last)  : 2001/12/01
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Park, Bum Soo 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ�.
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
Call LoadBasisGlobalInf
Call HideStatusWnd

On Error Resume Next

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter ���� 
Dim	rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim strConsumedDtFrom, strConsumedDtTo, strItemCd, strWcCd, strProdtOrderNo, strResourceCd, strResourceGroupCd

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgStrPrevKey1	
Dim i

Err.Clear																'��: Protect system from crashing

	Redim UNISqlId(5)
	Redim UNIValue(5, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000san"
	UNISqlId(2) = "180000sab"
	UNISqlId(3) = "180000sac"
	UNISqlId(4) = "180000sae"
	UNISqlID(5)	= "180000sal"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(4, 1) = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	UNIValue(5, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(5, 1) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5, rs6)
	Set ADF = Nothing
	
	'Call ServerMesgBox(UNIValue(5, 0) , vbInformation, I_MKSCRIPT)
	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbcr
		Response.Write "parent.frm1.txtPlantCd.focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbcr
		Response.Write "</Script>" & vbcr
	End If
	rs1.Close
	Set rs1 = Nothing
	
	'�̷¹�ȣ ��ȸ 
	If (rs6.EOF And rs6.BOF) Then
		rs6.Close
		Set rs6 = Nothing
		'Call ServerMesgBox("�̷¹�ȣ�� �������� �ʽ��ϴ�." , vbInformation, I_MKSCRIPT)
		Call DisplayMsgBox("189719", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtBatchRunNo.Focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End
	Else
	End If
	rs6.Close
	Set rs6 = Nothing

	' �ڿ��� Display
	If (rs2.EOF And rs2.BOF) Then
		rs2.Close
		Set rs2 = Nothing
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtResourceNm.value = """"" & vbcr
		Response.Write "</Script>" & vbcr
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtResourceNm.value = """ & ConvSPChars(rs2("description")) & """" & vbcr
		Response.Write "</Script>" & vbcr
	End If
	rs2.Close
	Set rs2 = Nothing

	' ǰ��� Display
	IF Request("txtItemCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbcr
			Response.Write "parent.frm1.txtItemCd.focus()" & vbcr
			Response.Write "</Script>"
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs3("ITEM_NM")) & """" & vbcr
			Response.Write "</Script>" & vbcr
		End If
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbcr
		Response.Write "</Script>" & vbcr
	End IF
	rs3.Close
	Set rs3 = Nothing

	' �۾���� Display
	IF Request("txtWcCd") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtWCNm.value = """"" & vbcr
			Response.Write "parent.frm1.txtWCCD.focus()" & vbcr
			Response.Write "</Script>" & vbcr
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtWCNm.value = """ & ConvSPChars(rs4("WC_NM")) & """" & vbcr
			Response.Write "</Script>" & vbcr
		End If
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtWCNm.value = """"" & vbcr
		Response.Write "</Script>" & vbcr
	End IF
	rs4.Close
	Set rs4 = Nothing
	
	' �ڿ��׷� �� Display
	IF Request("txtResourceGroupCd") <> "" Then
		If (rs5.EOF And rs5.BOF) Then
			Call DisplayMsgBox("181700", vbOKOnly, "", "", I_MKSCRIPT)
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtResourceGroupNm.value = """"" & vbcr
			Response.Write "parent.frm1.txtResourceGroupCd.focus()" & vbcr
			Response.Write "</Script>" & vbcr
			Response.End
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "parent.frm1.txtResourceGroupNm.value = """ & ConvSPChars(rs5("DESCRIPTION")) & """" & vbcr
			Response.Write "</Script>" & vbcr
		End IF
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "parent.frm1.txtResourceGroupNm.value = """"" & vbcr
		Response.Write "</Script>" & vbcr
	End IF
	rs5.Close
	Set rs5 = Nothing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "p4711mb4"

	IF Trim(Request("txtProdtOrderNo")) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	END IF
		
	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtConsumedDtFrom")) = "" Then
	   strConsumedDtFrom = "|"
	ELSE
	   strConsumedDtFrom = " " & FilterVar(UNIConvDate(Request("txtConsumedDtFrom")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtConsumedDtTo")) = "" Then
	   strConsumedDtTo = "|"
	ELSE
	   strConsumedDtTo = " " & FilterVar(UNIConvDate(Request("txtConsumedDtTo")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtResourceCd")) = "" Then
	   strResourceCd = "|"
	ELSE
	   strResourceCd = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtResourceGroupCd")) = "" Then
	   strResourceGroupCd = "|"
	ELSE
	   strResourceGroupCd = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	UNIValue(0, 3) = strProdtOrderNo
	UNIValue(0, 4) = strResourceCd
	UNIValue(0, 5) = strConsumedDtFrom
	UNIValue(0, 6) = strConsumedDtTo
	UNIValue(0, 7) = strResourceGroupCd
	UNIValue(0, 8) = strItemCd	
	UNIValue(0, 9) = strWcCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	Set ADF = Nothing
	
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
		rs0.Close
		Set rs0 = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_type"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_group_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("resource_group_nm"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("valid_from_dt"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("valid_to_dt"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData1
	.ggoSpread.SSShowDataByClip iTotalStr
	
	.lgStrPrevKey1 = "<%=Trim(rs0("resource_cd"))%>"

<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
