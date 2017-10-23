<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4110mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/01/23
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Chen Jae Hyun
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2	'DBAgent Parameter ���� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i
Dim j

Call HideStatusWnd

On Error Resume Next

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(2)
	Redim UNIValue(2, 2)
	
	UNISqlId(0) = "189702saa"
	UNISqlId(1) = "189702sab"
	UNISqlId(2) = "189702sae"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")
	
	UNIValue(1, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 1) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")

	UNIValue(2, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIvalue(2, 1) = FilterVar("C", "''", "S") 
	UNIValue(2, 2) = FilterVar(UCase(Request("txtPlanOrderNo")), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
      
	If (rs0.EOF And rs0.BOF) and (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing			
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>

Dim LngMaxRow1
Dim LngMaxRow2
Dim strData1
Dim strData2
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
    IF .parent.CompareDateByFormat("<%=UNIDateClientFormat(rs2("end_dt"))%>",.frm1.txtDlvyDt.text,"","","970025",gDateFormat,gComDateType,False) = False And "<%=rs2("push_flg")%>" = "Y" Then  
		.frm1.txtDlvyDt.text = "<%= UNIDateClientFormat(rs2("end_dt")) %>"
	End If

	LngMaxRow1 = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	LngMaxRow2 = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
%>		
		ReDim TmpBuffer(<%=rs0.RecordCount-1%>)
<%		
		For i=0 to rs0.RecordCount-1
%>	
			strData1 = ""		
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"			
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"			
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("start_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=UNIDateClientFormat(rs0("due_dt"))%>"
			strData1 = strData1 & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("plan_qty"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData1 = strData1 & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
			strData1 = strData1 & Chr(11) & LngMaxRow + <%=i%>
			strData1 = strData1 & Chr(11) & Chr(12)

			TmpBuffer(<%=i%>) = strData1
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
<%	
	End If
%>	
		
<%  
	If Not(rs1.EOF And rs1.BOF) Then
%>
		ReDim TmpBuffer(<%=rs1.RecordCount-1%>)
<%	
		For j=0 to rs1.RecordCount-1
%>			
			strData2 = ""
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_cd"))%>"
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("item_nm"))%>"			
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("spec"))%>"			
			strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("start_dt"))%>"
			strData2 = strData2 & Chr(11) & "<%=UNIDateClientFormat(rs1("due_dt"))%>"
			strData2 = strData2 & Chr(11) & "<%=UniConvNumberDBToCompany(rs1("plan_qty"),  ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData2 = strData2 & Chr(11) & "<%=ConvSPChars(rs1("tracking_no"))%>"
			strData2 = strData2 & Chr(11) & LngMaxRow + <%=j%>
			strData2 = strData2 & Chr(11) & Chr(12)
			
			TmpBuffer(<%=j%>) = strData2
<%		
			rs1.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		

<%
	End If

		rs0.Close
		Set rs0 = Nothing

		rs1.Close
		Set rs1 = Nothing
%>
	.DbQueryOk

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
