<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4212rb1.asp
'*  4. Program Name         : List Production Order Header
'*  5. Program Desc         : 
'*  6. Modified date(First) : Park , Bumsoo
'*  7. Modified date(Last)  : 2002/12/20
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================%>
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
Dim rs0, rs1, rs2, rs3						'DBAgent Parameter ���� 
Dim strMode									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Dim strItemCd
Dim StrProdOrderNo
Dim StrWcCd
Dim StrTrackingNo
Dim strOrderType
Dim strFlag
Dim strSlCd

On Error Resume Next
Err.Clear

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000saf"
	UNISqlId(2) = "180000sad"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtSlCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3)

	%>
	<Script Language=vbscript>
		parent.txtPlantNm.value = ""
		parent.txtItemNm.value = ""
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

	' ǰ��� Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			strFlag = "ERROR_ITEM"
			%>
			<Script Language=vbscript>
				parent.txtItemNm.value = ""
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
				parent.txtItemSpec.value = "<%=ConvSPChars(rs2("SPEC"))%>"
				parent.txtSafetyStock.value = "<%=ConvSPChars(rs2("SS_QTY"))%>"
				
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	Else
		%>
		<Script Language=vbscript>
			parent.txtItemNm.value = ""
		</Script>	
		<%
	End IF

	' â��� Display
	IF Request("txtSlCd") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
			strFlag = "ERROR_SLCD"
			%>
			<Script Language=vbscript>
				parent.txtSLNm.value = ""
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.txtSLNm.value = "<%=ConvSPChars(rs3("SL_NM"))%>"
			</Script>	
			<%
			rs3.Close
			Set rs3 = Nothing
		End If
	Else
		%>
		<Script Language=vbscript>
			parent.txtSLNm.value = ""
		</Script>	
		<%
	End IF

	If strFlag <> "" Then
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtPlantNm.Focus()
			</Script>	
			<%
			Response.End
		ElseIf strFlag = "ERROR_ITEM" Then
			Call DisplayMsgBox("122600", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtItemNm.Focus()
			</Script>	
			<%
			Response.End
		ElseIf strFlag = "ERROR_SLCD" Then
			Call DisplayMsgBox("125700", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.txtSLCd.Focus()
			</Script>	
			<%
			Response.End		
		End If
	End IF

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "P4212RB1"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtSlCd") = "" Then
		strSlCd = "|"
	Else
		StrSlCd = FilterVar(UCase(Request("txtSlCd")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strItemCd 
	UNIValue(0, 3) = strSlCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
			Parent.DbQueryNotOk()
		</Script>
		<%
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .vspdData1.MaxRows										'Save previous Maxrow
<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sl_cd"))%>"															'ǰ�� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sl_nm"))%>"															'ǰ��� 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("good_on_hand_qty"),ggQty.DecPoint,0)%>"							'�������� 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("schd_rcpt_qty"),ggQty.DecPoint,0)%>"							'�������� 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("schd_issue_qty"),ggQty.DecPoint,0)%>"							'�������� 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("avail_qty"),ggQty.DecPoint,0)%>"							'�������� 
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
				
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .vspdData1
	.ggoSpread.SSShowDataByClip iTotalStr
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.hItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.hSlCd.value		= "<%=ConvSPChars(Request("txtSlCd"))%>"
		
	.DbQueryOk()

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
