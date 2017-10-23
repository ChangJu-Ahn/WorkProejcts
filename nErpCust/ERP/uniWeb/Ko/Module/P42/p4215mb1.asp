<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4215mb1.asp
'*  4. Program Name         : List Order Document
'*  5. Program Desc         : 189400saa
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2002/07/10
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Chen, Jae Hyun
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
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")

On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1	'DBAgent Parameter ���� 
Dim strQryMode

Const C_SHEETMAXROWS = 30

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgStrPrevKey									' ���� �� 
Dim strOprCd
Dim strFlag
Dim LngRow


Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

	lgStrPrevKey = Request("lgStrPrevKey")
	
	Err.Clear										'��: Protect system from crashing
	
		'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)

	UNISqlId(0) = "180000saa"

	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

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

	End IF
	
    
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=====================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)
	
	UNISqlId(0) = "189400saa"
	
	IF Request("txtOprCd") = "" Then
		strOprCd = "|"
	Else
		StrOprCd = FilterVar(UCase(Request("txtOprCd")), "''", "S")
	End IF 		
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	Select Case strQryMode
	Case CStr(OPMD_CMODE)
		UNIValue(0, 3) = strOprCd		
	Case CStr(OPMD_UMODE)
		UNIValue(0, 3) = FilterVar(UCase(lgStrPrevKey), "''", "S")
	End Select

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		
		rs0.Close
		Set rs0 = Nothing
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strTemp
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount-1%>)	
<%  
    For LngRow= 0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("JOB_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("JOB_NM"))%>"
			
		If strMatchFlag = "N" Then 
			strData = strData & Chr(11) & ""
		End If			
			
		strTemp = "<%=ConvSPChars(rs0("Inside_Flg"))%>" 				
		If  strTemp = "Y" Then
			strData = strData & Chr(11) & "�系"
		ElseIf strTemp = "N" Then
			strData = strData & Chr(11) & "����"
		Else
			strData = strData & Chr(11) & ""		
		End If
			
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("status_nm"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("dtl_plan_start_dt")) %>"   		 
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("dtl_compt_start_dt"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("dtl_release_dt"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("GOOD_QTY_IN_ORDER_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("BAD_QTY_IN_ORDER_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_UNIT"))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(replace(rs0("document"),Chr(13) &Chr(10),chr(7)))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"		
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prodt_order_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_START_DT"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("PLAN_COMPT_DT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(replace(rs0("document"),Chr(13) &Chr(10),chr(7)))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>			
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=LngRow%>) = strData
			
<%		
		rs0.MoveNext
	Next
%>	
	iTotalStr = Join(TmpBuffer, "")	
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.lgStrPrevKey = "<%=ConvSPChars(rs0("OPR_NO"))%>"

	.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hProdOrderNo.value = "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hOprCd.value = "<%=ConvSPChars(Request("txtOprCd"))%>"

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
