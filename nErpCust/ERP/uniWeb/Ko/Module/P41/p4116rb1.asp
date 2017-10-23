<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4116rb1.asp
'*  4. Program Name         : List Conversion History
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002-04-25
'*  7. Modified date(Last)  : 2002/12/20
'*  8. Modifier (First)     : Park , Bumsoo
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=====================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1								'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim strQryMode
Dim i

Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

Dim StrProdOrderNo
Dim StrPRNo
Dim strFlag

On Error Resume Next
Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 1)

	UNISqlId(0) = "180000saa"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

	%>
	<Script Language=vbscript>
		parent.txtPlantNm.value = ""
	</Script>	
	<%

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		%>
		<Script Language=vbscript>
			parent.txtPlantNm.value = ""
		</Script>	
		<%    	
	Else
		%>
		<Script Language=vbscript>
			parent.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If


	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtPlantNm.Focus()
			</Script>	
			<%
			Response.End
		End If
	End IF

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	If Trim(Request("txtProdtOrderNo")) = "" Then
				StrProdOrderNo = "|"
	Else
		StrProdOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	End If

	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			StrPRNo = "|"
		Case CStr(OPMD_UMODE)
			If Trim(Request("lgStrPrevKey")) = "" Then
				StrPRNo = "|"
			Else
				StrPRNo = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
			End If
	End Select		

	UNISqlId(0) = "P4116RB1"
	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = StrProdOrderNo
	UNIValue(0, 3) = StrPRNo
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
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
	LngMaxRow = .vspdData.MaxRows										'Save previous Maxrow
		
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PR_NO"))%>"								'PR
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REQ_DT"))%>"						'��û�� 
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("DLVY_DT"))%>"					'�ʿ���	
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REQ_QTY"),ggQty.DecPoint,0)%>"	'��û���� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_UNIT"))%>"							'��û���� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG"))%>"							'�������� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"								'�԰�â�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_PRSN"))%>"							'��û�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_DEPT"))%>"							'��û�μ� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_STATUS"))%>"						'�������û��� 
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("INSRT_DT"))%>"					'��ȯ�� 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"								'��� 
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey = "<%=ConvSPChars(rs0("PR_NO"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.hProdOrderNo.value		= "<%=ConvSPChars(Request("txtProdtOrderNo"))%>"
	.DbQueryOk()

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
