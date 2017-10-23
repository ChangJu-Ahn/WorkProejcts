<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4312mb1.asp
'*  4. Program Name			: List Goods Issue (Query)
'*  5. Program Desc			: List Goods Issue (Called By Cancel Goods Issue )
'*  6. Comproxy List		: +P32119LookUpProdOrderHeader
'*                            +189660sab
'*  7. Modified date(First)	: 2000/03/28
'*  8. Modified date(Last)	: 2002/11/22
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: khk
'* 11. Comment		:
'**********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 

Call HideStatusWnd

On Error Resume Next

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter ���� 
Dim	rs0, rs1, rs2, rs3
Dim strMode													'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim i

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================

strMode = Request("txtMode")																'�� : ���� ���¸� ���� 

Dim strItemCd
Dim StrProdOrderNo
Dim strFlag

Err.Clear                                                      							'��: Protect system from crashing
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtCompntCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)

	' Plant �� Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
		</Script>	
		<%    	
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
	IF Request("txtCompntCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
			%>
			<Script Language=vbscript>
				parent.frm1.txtCompntNm.value = ""
			</Script>	
			<%
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtCompntNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
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
			Set gActiveElement = document.activeElement
			Response.End	
		End If
	End IF	
	
	Redim UNISqlId(1)
	Redim UNIValue(1, 7)

	UNISqlId(0) = "p4312mb1h"
	UNISqlId(1) = "189660sab"
	
	StrProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")

	IF Request("txtCompntCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtCompntCd")), "''", "S")
	End IF

	UNIValue(0, 0) = StrProdOrderNo

	UNIValue(1, 0) = "^"
	UNIValue(1, 1) = "" & FilterVar("PI", "''", "S") & ""
	UNIValue(1, 2) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 3) = StrProdOrderNo
	UNIValue(1, 4) = strItemCd 
	UNIValue(1, 5) = "" & FilterVar("M", "''", "S") & " "
	UNIValue(1, 6) = "" & FilterVar("N", "''", "S") & " "
	UNIValue(1, 7) = "" & FilterVar("CL", "''", "S") & ""

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs3, rs0)

	If (rs3.EOF And rs3.BOF) Then
		Call DisplayMsgBox("189200", vbOKOnly, "", "", I_MKSCRIPT)
		rs3.Close
		Set rs3 = Nothing
		%>
		<Script Language=vbscript>
		parent.frm1.txtProdOrderNo.Focus()
		Set gActiveElement = parent.document.activeElement
		</Script>	
		<%
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	Else
		%>
		<Script Language=vbscript>
			With parent.frm1

				.txtItemCd.value		= "<%=ConvSPChars(rs3("Item_Cd"))%>"
				.txtItemNm.value		= "<%=ConvSPChars(rs3("Item_Nm"))%>"
				.txtOrderQty.value		= "<%=UniNumClientFormat(rs3("Prodt_Order_Qty"),ggQty.DecPoint,0)%>"
				.txtPlndStartDt.text	= "<%=UNIDateClientFormat(rs3("Plan_Start_Dt"))%>"
				.txtPlndComptDt.text	= "<%=UNIDateClientFormat(rs3("Plan_Compt_Dt"))%>"
				.txtProdQty.Value		= "<%=UniNumClientFormat(rs3("Prod_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				.txtInspQty.Value		= "<%=UniNumClientFormat(rs3("Good_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"
				.txtRcptQty.Value 		= "<%=UniNumClientFormat(rs3("Rcpt_Qty_In_Order_Unit"),ggQty.DecPoint,0)%>"

			End With
		</Script>
		<%   	
	End If

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		%>
		<Script Language=vbscript>
		parent.frm1.txtProdOrderNo.Focus()
		Set gActiveElement = parent.document.activeElement
		parent.HeaderQueryOk
		</Script>	
		<%
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
		
<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>			
		ReDim TmpBuffer(1000)
<%
		For i=0 to rs0.RecordCount-1 
			if i=1000 then exit for
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"						'ǰ�� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"						'ǰ��� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"							'�԰� 
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("POS_DT"))%>"					'�����			
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("QTY"),ggQty.DecPoint,0)%>"	'�԰���� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>"						'���� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"							'�۾��� 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>"							'Lot No.
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>"						'Lot Sub No.
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_DOCUMENT_NO"))%>"				'��ǥ��ȣ 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REQ_NO"))%>"							'Requirement No.
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SEQ_NO"))%>"							'����			(Sequence No.)
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SUB_SEQ_NO"))%>"						'����			(Sub Sequence No.)
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>"							'���â��		(Storage Location)
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DOCUMENT_YEAR"))%>"					'��ǥ�߻��⵵	(Document Year)
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.DbQueryOk																'��: ��ȸ ������ ������� 

End With

</Script>	
<%
Set ADF = Nothing															'��: ActiveX Data Factory Object Nothing
%>
