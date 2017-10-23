<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4315mb1
'*  4. Program Name         : Query Component Reservation
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/15
'*  7. Modified date(Last)  : 2002/12/17
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Chen Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4

Const C_SHEETMAXROWS_D = 100

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strQryMode

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim strFlag
Dim strFromDt
Dim strToDt
Dim strProdOrderNo
Dim strPlantCd
Dim strItemCd1
Dim strItemCd2
Dim strConWcCd
Dim strItemAcct
Dim i

Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sab"
	UNISqlId(3) = "180000sac"
	
	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(1, 0) = FilterVar(Request("txtItemCd1"), "''", "S")
	UNIValue(2, 0) = FilterVar(Request("txtItemCd2"),"" & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S") & "","S")
	UNIValue(3, 0) = FilterVar(Request("txtConWcCd"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
		parent.frm1.txtItemNm2.value = ""
		parent.frm1.txtConWcNm.value = ""
	</Script>	
	<%
	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If

	' ǰ��� Display
	IF Trim(Request("txtItemCd1")) <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF

	' ǰ��� Display
	IF Trim(Request("txtItemCd2")) <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm2.value = "<%=ConvSPChars(rs3("ITEM_NM"))%>"
			</Script>	
			<%
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF
		
	' �۾���� Display
	IF Trim(Request("txtConWcCd")) <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			If strFlag <> "ERROR_PLANT" Then
				strFlag = "ERROR_WCCD"
			End If	
			
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtConWcNm.value = "<%=ConvSPChars(rs4("WC_NM"))%>"
			</Script>	
			<%
			rs4.Close
			Set rs4 = Nothing
		End If
	End IF
	
	If strFlag <> "" Then	
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
				parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Response.End	
		End If
		If strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtConWcCd.Focus()
			</Script>	
			<%
			Response.End
		End If
	End IF
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)
	
	strPlantCd	= FilterVar(Request("txtPlantCd"), "''", "S")
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			If Request("txtItemCd1") = "" Then
				strItemCd1	= "|"
			Else
				strItemCd1	= FilterVar(Request("txtItemCd1"), "''", "S")
			End If
		Case CStr(OPMD_UMODE) 
			
			strItemCd1	= FilterVar(Request("lgStrPrevKey1"), "''", "S")
			
	End Select
	
	If Request("txtItemCd2") = "" Then
		strItemCd2	= "|"
	Else
		strItemCd2	= FilterVar(Request("txtItemCd2"),"" & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S") & "","S")
	End If
	
	If Request("txtItemAcct") = "" Then
		strItemAcct	= "|"
	Else
		strItemAcct	= FilterVar(Request("txtItemAcct"), "''", "S")
	End If
		 
    If Request("txtConWcCd") = "" Then
		strConWcCd	= "|"
	Else
		strConWcCd	= FilterVar(Request("txtConWcCd"), "''", "S")
	End If
    
	If Request("txtFromDt") = "" Then
		strFromDt	= "|"
	Else
		strFromDt	= FilterVar(UniConvDate(Request("txtFromDt")), "''", "S")
	End If
	
	If Request("txtToDt") = "" Then
		strToDt	= "|"
	Else
		strToDt	= FilterVar(UniConvDate(Request("txtToDt")), "''", "S")
	End If
	
	UNISqlId(0) = "P4315MB1"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strItemCd1
	UNIValue(0, 3) = strItemCd2
	UNIValue(0, 4) = strFromDt
	UNIValue(0, 5) = strToDt
	UNIValue(0, 6) = strItemAcct
	UNIValue(0, 7) = strConWcCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Set rs0 = Nothing					
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
		  
%>
<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow		
<%  
	If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
	End If

    For i = 0 to rs0.RecordCount-1
		If i < C_SHEETMAXROWS_D Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
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
		
		
		.lgStrPrevKey1 = "<%=Trim(rs0("item_cd"))%>"
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hConWcCd.value = "<%=ConvSPChars(Request("txtConWcCd"))%>"
		.frm1.hFromDate.value = "<%=Request("txtFromDt")%>"
		.frm1.hToDate.value = "<%=Request("txtToDt")%>"
		.frm1.hItemCd1.value = "<%=Request("txtItemCd1")%>"
		.frm1.hItemCd2.value = "<%=Request("txtItemCd2")%>"
		.frm1.hItemAcct.value = "<%=Request("txtItemAcct")%>"		
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
